import contextlib
import os
import shutil
import tempfile
import zipfile
from collections.abc import Iterable, Mapping, Sequence
from concurrent.futures import ProcessPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from lxml import etree

from .constants import (
    APP_XML_PATH,
    CALC_CHAIN_CT,
    CALC_CHAIN_PATH,
    CALC_CHAIN_REL_TYPE,
    NS_CT,
    NS_EXT_PROPS,
    NS_MAIN,
    NS_REL_DOC,
    NS_REL_PKG,
    NS_VT,
    TABLE_REL_TYPE,
    TAG_AUTO_FILTER,
    TAG_CALC_COL_FORMULA,
    TAG_DEFINED_NAME,
    TAG_DEFINED_NAMES,
    TAG_FILTER_COLUMN,
    TAG_TABLE_COLUMN,
    TAG_TABLE_COLUMNS,
    TAG_TOTALS_ROW_FORMULA,
    WORKSHEET_CT,
    WORKSHEET_REL_TYPE,
    PatchWorkerFn,
)
from .models import (
    NewSheetSpec,
    SheetEditPlan,
    SheetMoveSpec,
    SheetStructureMutation,
    WorksheetTaskPayload,
)
from .patcher import create_sheet_worker
from .reader import WORKBOOK_RELS_PATH, read_sheet_infos
from .utils.coords import parse_range
from .utils.names import validate_sheet_name
from .utils.shifting import AxisShiftOp, rewrite_formula, shift_index, shift_range_str
from .utils.xml_text import parse_xml
from .utils.zip_io import (
    clone_zip_info,
    copy_disk_file_to_zip,
    new_zip_info,
    read_relationships,
    resolve_zip_path,
    stream_zip_entry,
    zip_entry_contains_formulas,
)


@dataclass(frozen=True, slots=True)
class _AllocatedNewSheet:
    name: str
    sheet_id: str
    rel_id: str
    zip_path: str
    rel_target: str
    position: int | None


def _worker_entrypoint(
    patch_fn: PatchWorkerFn,
    src_path: str,
    sheet_zip_path: str,
    edits: Any,
    temp_dir: str,
) -> tuple[str, str]:
    """
    Top-level picklable wrapper executed in worker processes.
    Writes XML to a temp file and returns (sheet_zip_path, temp_file_path).
    """
    with tempfile.NamedTemporaryFile(
        prefix="sheet_patch_", suffix=".xml", dir=temp_dir, delete=False
    ) as tmp_f:
        tmp_xml_path = Path(tmp_f.name)
    try:
        patch_fn(src_path, sheet_zip_path, edits, str(tmp_xml_path))
        return sheet_zip_path, str(tmp_xml_path)
    except Exception:
        tmp_xml_path.unlink(missing_ok=True)
        raise


class PackageAssembler:
    """
    Coordinates multiprocessing across dirty and newly created sheets, performs
    sheet structural operations (create, delete, reorder, row/col shifting,
    Excel Table resizing, and formula rewriting), and atomically writes the
    resulting .xlsx/.xlsm archive (in-place by default, or to dst_path).
    """

    # Rough serialized size of one generated <c> element, used to estimate new-sheet work
    _EST_BYTES_PER_NEW_CELL = 50

    def __init__(
        self,
        src_path: str | Path,
        dst_path: str | Path | None = None,
        max_workers: int | None = None,
        chunk_size: int = 1024 * 1024,  # 1 MB streaming buffer
        compresslevel: int = 1,
        parallel_threshold_bytes: int = 512 * 1024,
    ) -> None:
        self.src_path = Path(src_path)
        self.dst_path = Path(dst_path) if dst_path is not None else self.src_path
        self.max_workers = max_workers or os.process_cpu_count() or 1
        self.chunk_size = chunk_size
        self.compresslevel = compresslevel
        self.parallel_threshold_bytes = parallel_threshold_bytes

    def build(
        self,
        dirty_edits: Mapping[str, Any] | None = None,
        patch_fn: PatchWorkerFn | None = None,
        *,
        new_sheets: Sequence[NewSheetSpec] = (),
        delete_sheets: Iterable[str] = (),
        move_sheets: Sequence[SheetMoveSpec] = (),
        shift_ops: Mapping[str, tuple[Sequence[AxisShiftOp], Sequence[AxisShiftOp]]] | None = None,
        update_formulas: bool = True,
        force_recalc: bool = True,
    ) -> None:
        """
        Applies cell edits, row/column shifts, table resizes, formula rewrites,
        and worksheet structural mutations atomically.
        """
        dirty_map = dict(dirty_edits) if dirty_edits else {}
        delete_set = set(delete_sheets)
        normalized_shifts: dict[str, tuple[tuple[AxisShiftOp, ...], tuple[AxisShiftOp, ...]]] = {
            s_name: (tuple(r_ops), tuple(c_ops))
            for s_name, (r_ops, c_ops) in (shift_ops or {}).items()
            if r_ops or c_ops
        }

        if (
            not dirty_map
            and not new_sheets
            and not delete_set
            and not move_sheets
            and not normalized_shifts
        ):
            if self.src_path.resolve() != self.dst_path.resolve():
                shutil.copy2(self.src_path, self.dst_path)
            return

        with tempfile.TemporaryDirectory(prefix="xlsx_build_") as temp_dir:
            (
                allocated_new,
                deleted_zip_paths,
                surviving_sheets,
                table_shifts_by_zip,
                deleted_extra_parts,
            ) = self._plan_workbook_mutations(
                new_sheets=new_sheets,
                delete_set=delete_set,
                move_sheets=move_sheets,
                shift_ops=normalized_shifts,
            )

            # Exclude any sheets scheduled for deletion from patching
            effective_dirty: dict[str, Any] = {
                path: edits
                for path, edits in dirty_map.items()
                if path not in deleted_zip_paths
            }

            # Attach shift/formula payloads to locally shifted sheets and any
            # surviving sheets that actually contain '<f' formula tags.
            has_global_formula_shifts = bool(normalized_shifts or delete_set)
            if has_global_formula_shifts:
                deleted_tuple = tuple(sorted(delete_set))
                with zipfile.ZipFile(self.src_path, "r") as zin:
                    for s_name, s_zip_path in surviving_sheets.items():
                        r_ops, c_ops = normalized_shifts.get(s_name, ((), ()))
                        has_local = bool(r_ops or c_ops)
                        if not has_local and not update_formulas:
                            continue
                        if not has_local and not zip_entry_contains_formulas(
                            zin, s_zip_path, self.chunk_size
                        ):
                            continue

                        mutation = SheetStructureMutation(
                            sheet_name=s_name,
                            row_ops=r_ops,
                            col_ops=c_ops,
                            all_shifts=normalized_shifts if update_formulas else {},
                            deleted_sheets=deleted_tuple if update_formulas else (),
                        )
                        existing_item = effective_dirty.get(s_zip_path)
                        plan = (
                            existing_item.plan
                            if isinstance(existing_item, WorksheetTaskPayload)
                            else (existing_item or SheetEditPlan.empty())
                        )
                        effective_dirty[s_zip_path] = WorksheetTaskPayload(
                            plan=plan,
                            mutation=mutation,
                        )

            patched_files = self._run_sheet_workers(
                dirty_edits=effective_dirty,
                patch_fn=patch_fn,
                new_sheets=new_sheets,
                allocated_new=allocated_new,
                temp_dir=temp_dir,
            )

            self._assemble_zip(
                patched_files=patched_files,
                allocated_new=allocated_new,
                deleted_zip_paths=deleted_zip_paths,
                deleted_extra_parts=deleted_extra_parts,
                table_shifts_by_zip=table_shifts_by_zip,
                delete_sheet_names=delete_set,
                move_sheets=move_sheets,
                shift_ops=normalized_shifts if update_formulas else {},
                update_formulas=update_formulas,
                strip_calc_chain=force_recalc,
            )

    def _plan_workbook_mutations(
        self,
        new_sheets: Sequence[NewSheetSpec],
        delete_set: set[str],
        move_sheets: Sequence[SheetMoveSpec],
        shift_ops: Mapping[str, tuple[tuple[AxisShiftOp, ...], tuple[AxisShiftOp, ...]]],
    ) -> tuple[
        list[_AllocatedNewSheet],
        set[str],
        dict[str, str],
        dict[str, tuple[tuple[AxisShiftOp, ...], tuple[AxisShiftOp, ...]]],
        set[str],
    ]:
        """
        Validates sheet operations, resolves Excel Table (xl/tables/tableN.xml)
        targets for shifted sheets, and allocates non-conflicting sheetN.xml paths,
        sheetId integers, and rId strings for new worksheets.
        """
        with zipfile.ZipFile(self.src_path, "r") as zin:
            existing_entries = set(zin.namelist())
            wb_rels = read_relationships(zin, WORKBOOK_RELS_PATH)
            sheet_infos = read_sheet_infos(zin, wb_rels)

            existing_sheets_by_name = {
                name: info.zip_path for name, info in sheet_infos.items()
            }
            visible_names = {
                name for name, info in sheet_infos.items() if info.state == "visible"
            }
            max_sheet_id = max(
                (int(info.sheet_id) for info in sheet_infos.values() if info.sheet_id.isdigit()),
                default=0,
            )

            # Validate deletions
            unknown_deletes = delete_set - existing_sheets_by_name.keys()
            if unknown_deletes:
                raise KeyError(
                    f"Cannot delete unknown worksheet(s): {sorted(unknown_deletes)}"
                )

            remaining_visible = len(visible_names - delete_set) + len(new_sheets)
            if remaining_visible < 1:
                raise ValueError("Workbook must retain at least one visible worksheet.")

            deleted_zip_paths = {existing_sheets_by_name[name] for name in delete_set}
            surviving_sheets = {
                name: zip_path
                for name, zip_path in existing_sheets_by_name.items()
                if name not in delete_set
            }

            # Validate shift targets
            unknown_shifts = set(shift_ops.keys()) - set(existing_sheets_by_name.keys())
            if unknown_shifts:
                raise KeyError(
                    f"Cannot shift rows/cols on unknown worksheet(s): {sorted(unknown_shifts)}"
                )

            # Map Excel Tables (xl/tables/tableN.xml) belonging to deleted or shifted sheets
            table_shifts_by_zip: dict[
                str, tuple[tuple[AxisShiftOp, ...], tuple[AxisShiftOp, ...]]
            ] = {}
            deleted_extra_parts: set[str] = set()

            for s_name, s_zip_path in existing_sheets_by_name.items():
                s_rels_path = self._rels_path_for_sheet(s_zip_path)
                if s_rels_path not in existing_entries:
                    continue
                for rel in read_relationships(zin, s_rels_path):
                    if rel.type == TABLE_REL_TYPE and not rel.external:
                        table_zip = resolve_zip_path(s_zip_path, rel.target)
                        if s_name in delete_set:
                            deleted_extra_parts.add(table_zip)
                        elif s_name in shift_ops:
                            if table_zip in existing_entries:
                                self._validate_table_shift(
                                    zin.read(table_zip), *shift_ops[s_name], table_zip
                                )
                            table_shifts_by_zip[table_zip] = shift_ops[s_name]

            if deleted_zip_paths:
                deleted_extra_parts |= self._find_orphaned_parts(
                    zin, existing_entries, deleted_zip_paths
                )

        # Validate new sheet names against surviving + newly added names (Excel is case-insensitive)
        surviving_names = set(surviving_sheets.keys())
        taken_folded = {name.casefold() for name in surviving_names}
        seen_new_names: set[str] = set()
        for spec in new_sheets:
            validate_sheet_name(spec.name)
            folded = spec.name.casefold()
            if folded in taken_folded:
                raise ValueError(f"Duplicate worksheet name: {spec.name!r}")
            taken_folded.add(folded)
            seen_new_names.add(spec.name)

        # Validate move targets
        final_names = surviving_names | seen_new_names
        for move in move_sheets:
            if move.name not in final_names:
                raise KeyError(f"Cannot move unknown worksheet: {move.name!r}")

        # Allocate unique sheetN.xml, sheetId, and rId for each new sheet
        used_rids = {rel.id for rel in wb_rels}
        next_rid_num = 1
        next_sheet_file_num = 1
        next_sheet_id = max_sheet_id + 1

        allocated: list[_AllocatedNewSheet] = []
        for spec in new_sheets:
            while f"rId{next_rid_num}" in used_rids:
                next_rid_num += 1
            rel_id = f"rId{next_rid_num}"
            used_rids.add(rel_id)

            while f"xl/worksheets/sheet{next_sheet_file_num}.xml" in existing_entries:
                next_sheet_file_num += 1
            zip_path = f"xl/worksheets/sheet{next_sheet_file_num}.xml"
            rel_target = f"worksheets/sheet{next_sheet_file_num}.xml"
            existing_entries.add(zip_path)

            allocated.append(
                _AllocatedNewSheet(
                    name=spec.name,
                    sheet_id=str(next_sheet_id),
                    rel_id=rel_id,
                    zip_path=zip_path,
                    rel_target=rel_target,
                    position=spec.position,
                )
            )
            next_sheet_id += 1

        return (
            allocated,
            deleted_zip_paths,
            surviving_sheets,
            table_shifts_by_zip,
            deleted_extra_parts,
        )

    def _run_sheet_workers(
        self,
        dirty_edits: Mapping[str, Any],
        patch_fn: PatchWorkerFn | None,
        new_sheets: Sequence[NewSheetSpec],
        allocated_new: Sequence[_AllocatedNewSheet],
        temp_dir: str,
    ) -> dict[str, str]:
        """Dispatches dirty & new sheets to ProcessPoolExecutor or runs inline if small."""
        tasks: list[tuple[PatchWorkerFn, str, Any]] = []

        if dirty_edits:
            if patch_fn is None:
                raise ValueError("patch_fn is required when dirty_edits are provided.")
            for sheet_path, edits in dirty_edits.items():
                tasks.append((patch_fn, sheet_path, edits))

        for spec, alloc in zip(new_sheets, allocated_new, strict=True):
            tasks.append((create_sheet_worker, alloc.zip_path, spec.plan))

        if not tasks:
            return {}

        with zipfile.ZipFile(self.src_path, "r") as zin:
            total_bytes = sum(zin.getinfo(path).file_size for path in dirty_edits)
        total_bytes += self._EST_BYTES_PER_NEW_CELL * sum(
            len(batch.cells) for spec in new_sheets for batch in spec.plan.rows
        )

        # Process spawn costs ~100-300 ms on Windows; only worth it for 2+ sizeable tasks
        use_multiprocessing = (
            len(tasks) > 1
            and self.max_workers > 1
            and total_bytes >= self.parallel_threshold_bytes
        )

        patched_map: dict[str, str] = {}
        src_str = str(self.src_path)

        if not use_multiprocessing:
            for fn, sheet_path, edits in tasks:
                path, tmp_file = _worker_entrypoint(
                    fn, src_str, sheet_path, edits, temp_dir
                )
                patched_map[path] = tmp_file
            return patched_map

        workers = min(self.max_workers, len(tasks))
        with ProcessPoolExecutor(max_workers=workers) as pool:
            futures = {
                pool.submit(
                    _worker_entrypoint,
                    fn,
                    src_str,
                    sheet_path,
                    edits,
                    temp_dir,
                ): sheet_path
                for fn, sheet_path, edits in tasks
            }

            for future in as_completed(futures):
                sheet_path, tmp_file = future.result()
                patched_map[sheet_path] = tmp_file

        return patched_map

    def _assemble_zip(
        self,
        patched_files: Mapping[str, str],
        allocated_new: Sequence[_AllocatedNewSheet],
        deleted_zip_paths: set[str],
        deleted_extra_parts: set[str],
        table_shifts_by_zip: Mapping[
            str, tuple[tuple[AxisShiftOp, ...], tuple[AxisShiftOp, ...]]
        ],
        delete_sheet_names: set[str],
        move_sheets: Sequence[SheetMoveSpec],
        shift_ops: Mapping[str, tuple[tuple[AxisShiftOp, ...], tuple[AxisShiftOp, ...]]],
        update_formulas: bool,
        strip_calc_chain: bool,
    ) -> None:
        """
        Streams entries into the destination archive without buffering whole files,
        using fast deflate compression (compresslevel) across all streamed parts.
        """
        dst_dir = self.dst_path.parent
        dst_dir.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(
            prefix=".tmp_xlsx_", suffix=".xlsx", dir=dst_dir, delete=False
        ) as staging_f:
            staging_zip = Path(staging_f.name)

        all_deleted_parts = deleted_zip_paths | deleted_extra_parts
        deleted_rels_paths = {self._rels_path_for_sheet(p) for p in all_deleted_parts}
        update_app_xml = bool(allocated_new or delete_sheet_names or move_sheets)

        try:
            with (
                zipfile.ZipFile(self.src_path, "r") as zin,
                zipfile.ZipFile(
                    staging_zip,
                    "w",
                    compression=zipfile.ZIP_DEFLATED,
                    compresslevel=self.compresslevel,
                ) as zout,
            ):
                final_order: list[str] = []
                orig_sheet_names: list[str] = []
                if update_app_xml and APP_XML_PATH in zin.namelist():
                    wb_root = parse_xml(zin.read("xl/workbook.xml"))
                    orig_sheet_names = [
                        s.get("name", "") for s in wb_root.iter(f"{{{NS_MAIN}}}sheet")
                    ]
                    final_order = self._final_tab_order(
                        orig_sheet_names, allocated_new, delete_sheet_names, move_sheets
                    )

                for info in zin.infolist():
                    name = info.filename

                    if strip_calc_chain and name == CALC_CHAIN_PATH:
                        continue

                    if name in all_deleted_parts or name in deleted_rels_paths:
                        continue

                    if name == APP_XML_PATH and final_order:
                        zout.writestr(
                            clone_zip_info(info),
                            self._update_app_xml(
                                zin.read(name), orig_sheet_names, final_order
                            ),
                            compresslevel=self.compresslevel,
                        )
                        continue

                    if name in patched_files:
                        copy_disk_file_to_zip(
                            zout,
                            info,
                            patched_files[name],
                            self.chunk_size,
                            compresslevel=self.compresslevel,
                        )
                        continue

                    if name in table_shifts_by_zip:
                        r_ops, c_ops = table_shifts_by_zip[name]
                        updated_table = self._update_table_xml(
                            zin.read(name), r_ops, c_ops
                        )
                        zout.writestr(
                            clone_zip_info(info),
                            updated_table,
                            compresslevel=self.compresslevel,
                        )
                        continue

                    if name == "[Content_Types].xml":
                        updated_ct = self._update_content_types(
                            zin.read(name),
                            allocated_new=allocated_new,
                            deleted_zip_paths=all_deleted_parts,
                            strip_calc_chain=strip_calc_chain,
                        )
                        zout.writestr(
                            clone_zip_info(info),
                            updated_ct,
                            compresslevel=self.compresslevel,
                        )
                        continue

                    if name == "xl/_rels/workbook.xml.rels":
                        updated_rels = self._update_workbook_rels(
                            zin.read(name),
                            allocated_new=allocated_new,
                            deleted_zip_paths=deleted_zip_paths,
                            strip_calc_chain=strip_calc_chain,
                        )
                        zout.writestr(
                            clone_zip_info(info),
                            updated_rels,
                            compresslevel=self.compresslevel,
                        )
                        continue

                    if name == "xl/workbook.xml":
                        updated_wb = self._update_workbook_xml(
                            zin.read(name),
                            allocated_new=allocated_new,
                            delete_sheet_names=delete_sheet_names,
                            move_sheets=move_sheets,
                            shift_ops=shift_ops,
                            update_formulas=update_formulas,
                            force_recalc=strip_calc_chain,
                        )
                        zout.writestr(
                            clone_zip_info(info),
                            updated_wb,
                            compresslevel=self.compresslevel,
                        )
                        continue

                    stream_zip_entry(
                        zin,
                        zout,
                        info,
                        self.chunk_size,
                        compresslevel=self.compresslevel,
                    )

                # Append brand-new worksheet XML files
                for alloc in allocated_new:
                    disk_path = patched_files[alloc.zip_path]
                    copy_disk_file_to_zip(
                        zout,
                        new_zip_info(alloc.zip_path),
                        disk_path,
                        self.chunk_size,
                        compresslevel=self.compresslevel,
                    )

            # Temp files are created 0600; inherit the original file's permissions instead
            mode_source = self.dst_path if self.dst_path.exists() else self.src_path
            with contextlib.suppress(OSError):
                shutil.copymode(mode_source, staging_zip)
            staging_zip.replace(self.dst_path)
        except Exception:
            staging_zip.unlink(missing_ok=True)
            raise

    @staticmethod
    def _rels_path_for_sheet(sheet_zip_path: str) -> str:
        """Converts 'xl/worksheets/sheet1.xml' -> 'xl/worksheets/_rels/sheet1.xml.rels'."""
        prefix, _, filename = sheet_zip_path.rpartition("/")
        return f"{prefix}/_rels/{filename}.rels" if prefix else f"_rels/{filename}.rels"

    @staticmethod
    def _validate_table_shift(
        xml_bytes: bytes,
        row_ops: Sequence[AxisShiftOp],
        col_ops: Sequence[AxisShiftOp],
        table_zip: str,
    ) -> None:
        """Rejects shifts that would delete a table entirely or remove its header/totals row."""
        root = parse_xml(xml_bytes)
        ref = root.get("ref", "")
        if not ref:
            return
        name = root.get("displayName") or root.get("name") or table_zip
        if shift_range_str(ref, row_ops, col_ops) is None:
            raise ValueError(
                f"Operation would delete Excel Table {name!r} entirely; not supported."
            )
        r1, _, r2, _ = parse_range(ref)
        if root.get("headerRowCount", "1") != "0" and shift_index(r1, row_ops) is None:
            raise ValueError(
                f"Operation would delete the header row of Excel Table {name!r}; not supported."
            )
        if root.get("totalsRowCount", "0") not in ("", "0") and shift_index(r2, row_ops) is None:
            raise ValueError(
                f"Operation would delete the totals row of Excel Table {name!r}; not supported."
            )

    @staticmethod
    def _remap_table_columns(
        root: etree._Element,
        old_ref: str,
        new_ref: str,
        col_ops: Sequence[AxisShiftOp],
    ) -> None:
        """Drops deleted <tableColumn>s, adds placeholders for inserted ones, and remaps filter colIds."""
        cols_el = root.find(TAG_TABLE_COLUMNS)
        if cols_el is None:
            return
        _, old_c1, _, _ = parse_range(old_ref)
        _, new_c1, _, new_c2 = parse_range(new_ref)
        old_cols = cols_el.findall(TAG_TABLE_COLUMN)

        old_to_new_idx: dict[int, int] = {}
        col_by_new_pos: dict[int, etree._Element] = {}
        for i, tc in enumerate(old_cols):
            new_pos = shift_index(old_c1 + i, col_ops)
            if new_pos is not None:
                old_to_new_idx[i] = new_pos - new_c1
                col_by_new_pos[new_pos] = tc

        used_names = {(tc.get("name") or "").casefold() for tc in old_cols}
        next_id = 1 + max(
            (int(i) for tc in old_cols if (i := tc.get("id", "")).isdigit()), default=0
        )
        new_cols: list[etree._Element] = []
        for pos in range(new_c1, new_c2 + 1):
            tc = col_by_new_pos.get(pos)
            if tc is None:
                suffix = next_id
                while f"column{suffix}" in used_names:
                    suffix += 1
                used_names.add(f"column{suffix}")
                # Header cell text must match this name for Excel to accept the table
                tc = etree.Element(TAG_TABLE_COLUMN, id=str(next_id), name=f"Column{suffix}")
                next_id += 1
            new_cols.append(tc)
        cols_el[:] = new_cols
        cols_el.set("count", str(len(new_cols)))

        af = root.find(TAG_AUTO_FILTER)
        if af is not None:
            for fc in af.findall(TAG_FILTER_COLUMN):
                raw = fc.get("colId", "")
                new_idx = old_to_new_idx.get(int(raw)) if raw.isdigit() else None
                if new_idx is None:
                    af.remove(fc)
                else:
                    fc.set("colId", str(new_idx))

    @staticmethod
    def _update_table_xml(
        xml_bytes: bytes,
        row_ops: Sequence[AxisShiftOp],
        col_ops: Sequence[AxisShiftOp],
    ) -> bytes:
        """Resizes <table ref>, <autoFilter ref>, <tableColumns>, and column formulas on tableN.xml."""
        root = parse_xml(xml_bytes)
        ref = root.get("ref", "")
        if ref:
            new_ref = shift_range_str(ref, row_ops, col_ops)
            if new_ref is not None:
                root.set("ref", new_ref)
                if col_ops:
                    PackageAssembler._remap_table_columns(root, ref, new_ref, col_ops)

        af = root.find(TAG_AUTO_FILTER)
        if af is not None:
            af_ref = af.get("ref", "")
            if af_ref:
                new_af = shift_range_str(af_ref, row_ops, col_ops)
                if new_af is not None:
                    af.set("ref", new_af)

        # Only this sheet's local A1 refs are shifted; structured refs pass through untouched
        local_shifts = {"": (row_ops, col_ops)}
        for tc in root.iter(TAG_TABLE_COLUMN):
            for f_tag in (TAG_CALC_COL_FORMULA, TAG_TOTALS_ROW_FORMULA):
                f_el = tc.find(f_tag)
                if f_el is not None and f_el.text:
                    f_el.text = rewrite_formula(
                        f_el.text, current_sheet="", shifts_by_sheet=local_shifts
                    )

        return etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )

    @staticmethod
    def _update_content_types(
        xml_bytes: bytes,
        allocated_new: Sequence[_AllocatedNewSheet],
        deleted_zip_paths: set[str],
        strip_calc_chain: bool,
    ) -> bytes:
        root = parse_xml(xml_bytes)
        deleted_part_names = {f"/{p.lstrip('/')}" for p in deleted_zip_paths}

        for override in list(root.findall(f"{{{NS_CT}}}Override")):
            part_name = override.get("PartName", "")
            content_type = override.get("ContentType", "")

            if strip_calc_chain and (
                part_name == "/xl/calcChain.xml" or content_type == CALC_CHAIN_CT
            ):
                root.remove(override)
                continue

            if part_name in deleted_part_names:
                root.remove(override)

        for alloc in allocated_new:
            override = etree.SubElement(root, f"{{{NS_CT}}}Override")
            override.set("PartName", f"/{alloc.zip_path}")
            override.set("ContentType", WORKSHEET_CT)

        return etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )

    @staticmethod
    def _update_workbook_rels(
        xml_bytes: bytes,
        allocated_new: Sequence[_AllocatedNewSheet],
        deleted_zip_paths: set[str],
        strip_calc_chain: bool,
    ) -> bytes:
        root = parse_xml(xml_bytes)

        for rel in list(root.findall(f"{{{NS_REL_PKG}}}Relationship")):
            rel_type = rel.get("Type", "")
            target = rel.get("Target", "")

            if strip_calc_chain and rel_type == CALC_CHAIN_REL_TYPE:
                root.remove(rel)
                continue

            if target and resolve_zip_path("xl/workbook.xml", target) in deleted_zip_paths:
                root.remove(rel)

        for alloc in allocated_new:
            rel = etree.SubElement(root, f"{{{NS_REL_PKG}}}Relationship")
            rel.set("Id", alloc.rel_id)
            rel.set("Type", WORKSHEET_REL_TYPE)
            rel.set("Target", alloc.rel_target)

        return etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )

    @staticmethod
    def _final_tab_order(
        orig_names: Sequence[str],
        allocated_new: Sequence[_AllocatedNewSheet],
        delete_names: set[str],
        move_sheets: Sequence[SheetMoveSpec],
    ) -> list[str]:
        """Computes final sheet tab order: drop deletions, insert new sheets, then apply moves."""
        order = [n for n in orig_names if n not in delete_names]
        for alloc in allocated_new:
            if alloc.position is None:
                order.append(alloc.name)
            else:
                order.insert(max(0, min(alloc.position, len(order))), alloc.name)
        for move in move_sheets:
            if move.name not in order:
                continue
            order.remove(move.name)
            if move.position is None:
                order.append(move.name)
            else:
                order.insert(max(0, min(move.position, len(order))), move.name)
        return order

    @staticmethod
    def _owner_of_rels(rels_path: str) -> str:
        """Maps 'xl/worksheets/_rels/sheet1.xml.rels' -> 'xl/worksheets/sheet1.xml' ('' for package root)."""
        prefix, _, filename = rels_path.rpartition("/")
        owner_dir = prefix.removesuffix("_rels").rstrip("/")
        owner_name = filename.removesuffix(".rels")
        return f"{owner_dir}/{owner_name}" if owner_dir else owner_name

    @staticmethod
    def _find_orphaned_parts(
        zin: zipfile.ZipFile,
        existing_entries: set[str],
        deleted_sheet_paths: set[str],
    ) -> set[str]:
        """
        Mark-and-sweep over internal relationships: returns parts reachable from deleted
        sheets that are no longer reachable from the package root.
        """
        graph: dict[str, list[str]] = {}
        for entry in existing_entries:
            if not entry.endswith(".rels"):
                continue
            owner = PackageAssembler._owner_of_rels(entry)
            graph[owner] = [
                resolve_zip_path(owner, rel.target)
                for rel in read_relationships(zin, entry)
                if not rel.external
            ]

        def reachable(starts: Iterable[str], blocked: set[str]) -> set[str]:
            seen: set[str] = set()
            stack = list(starts)
            while stack:
                part = stack.pop()
                if part in seen or part in blocked:
                    continue
                seen.add(part)
                stack.extend(graph.get(part, ()))
            return seen

        candidates = reachable(deleted_sheet_paths, set()) - deleted_sheet_paths
        alive = reachable([""], deleted_sheet_paths)
        return {p for p in candidates - alive if p in existing_entries}

    @staticmethod
    def _update_app_xml(
        xml_bytes: bytes,
        orig_names: Sequence[str],
        final_order: Sequence[str],
    ) -> bytes:
        """
        Rewrites sheet-name groups in docProps/app.xml TitlesOfParts/HeadingPairs.
        Groups are matched by content (heading labels are localized); unknown layouts are left as-is.
        """
        root = parse_xml(xml_bytes)
        hp_vec = root.find(f"{{{NS_EXT_PROPS}}}HeadingPairs/{{{NS_VT}}}vector")
        tp_vec = root.find(f"{{{NS_EXT_PROPS}}}TitlesOfParts/{{{NS_VT}}}vector")
        if hp_vec is None or tp_vec is None:
            return xml_bytes

        variants = hp_vec.findall(f"{{{NS_VT}}}variant")
        titles = [t.text or "" for t in tp_vec.findall(f"{{{NS_VT}}}lpstr")]
        orig_set = set(orig_names)
        new_names = [n for n in final_order if n not in orig_set]

        new_titles: list[str] = []
        kept_variants: list[etree._Element] = []
        pos = 0
        for k in range(0, len(variants) - 1, 2):
            count_el = variants[k + 1].find(f"{{{NS_VT}}}i4")
            if count_el is None or not (count_el.text or "").isdigit():
                return xml_bytes
            count = int(count_el.text)
            group = titles[pos : pos + count]
            pos += count
            if group and all(g in orig_set for g in group):
                members = set(group) | set(new_names)
                new_names = []  # New worksheets join the first sheet-name group only
                group = [n for n in final_order if n in members]
                if not group:
                    continue
                count_el.text = str(len(group))
            kept_variants.extend((variants[k], variants[k + 1]))
            new_titles.extend(group)
        if pos != len(titles):
            return xml_bytes

        hp_vec[:] = kept_variants
        hp_vec.set("size", str(len(kept_variants)))
        lpstr_elems = []
        for title in new_titles:
            el = etree.Element(f"{{{NS_VT}}}lpstr")
            el.text = title
            lpstr_elems.append(el)
        tp_vec[:] = lpstr_elems
        tp_vec.set("size", str(len(lpstr_elems)))

        return etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )

    @staticmethod
    def _update_workbook_xml(
        xml_bytes: bytes,
        allocated_new: Sequence[_AllocatedNewSheet],
        delete_sheet_names: set[str],
        move_sheets: Sequence[SheetMoveSpec],
        shift_ops: Mapping[str, tuple[tuple[AxisShiftOp, ...], tuple[AxisShiftOp, ...]]],
        update_formulas: bool,
        force_recalc: bool,
    ) -> bytes:
        root = parse_xml(xml_bytes)
        sheets_elem = root.find(f"{{{NS_MAIN}}}sheets")
        if sheets_elem is None:
            sheets_elem = etree.SubElement(root, f"{{{NS_MAIN}}}sheets")

        # 1-3. Drop deleted sheets, insert new ones, and apply tab moves
        orig_nodes = sheets_elem.findall(f"{{{NS_MAIN}}}sheet")
        orig_order = [s.get("name", "") for s in orig_nodes]
        node_by_name = {s.get("name", ""): s for s in orig_nodes}
        for alloc in allocated_new:
            node_by_name[alloc.name] = etree.Element(
                f"{{{NS_MAIN}}}sheet",
                attrib={
                    "name": alloc.name,
                    "sheetId": alloc.sheet_id,
                    f"{{{NS_REL_DOC}}}id": alloc.rel_id,
                },
            )
        final_order = PackageAssembler._final_tab_order(
            orig_order, allocated_new, delete_sheet_names, move_sheets
        )
        sheet_nodes = [node_by_name[n] for n in final_order]

        # Replace <sheets> children in final tab order
        sheets_elem[:] = sheet_nodes
        new_index = {el.get("name", ""): i for i, el in enumerate(sheet_nodes)}

        def remap_index(raw: str | None) -> int | None:
            """Maps an original 0-based sheet index to its new index (None if deleted/invalid)."""
            if raw is None or not raw.isdigit() or int(raw) >= len(orig_order):
                return None
            return new_index.get(orig_order[int(raw)])

        # 4. Remap sheet-scoped <definedNames> and rewrite their formulas
        rewrite_names = update_formulas and bool(shift_ops or delete_sheet_names)
        defined_names = root.find(TAG_DEFINED_NAMES)
        if defined_names is not None:
            for dn in list(defined_names.findall(TAG_DEFINED_NAME)):
                cur_sheet = ""
                local_idx = dn.get("localSheetId")
                if local_idx is not None:
                    new_idx = remap_index(local_idx)
                    if new_idx is None:
                        defined_names.remove(dn)
                        continue
                    dn.set("localSheetId", str(new_idx))
                    cur_sheet = sheet_nodes[new_idx].get("name", "")
                if rewrite_names and dn.text:
                    dn.text = rewrite_formula(
                        dn.text,
                        current_sheet=cur_sheet,
                        shifts_by_sheet=shift_ops,
                        deleted_sheets=delete_sheet_names,
                    )
            if len(defined_names) == 0:
                root.remove(defined_names)

        # Keep activeTab/firstSheet pointing at the same sheet after reordering
        for wb_view in root.iter(f"{{{NS_MAIN}}}workbookView"):
            for attr in ("activeTab", "firstSheet"):
                if attr in wb_view.attrib:
                    new_idx = remap_index(wb_view.get(attr))
                    wb_view.set(attr, str(new_idx if new_idx is not None else 0))

        if force_recalc:
            calc_pr = root.find(f"{{{NS_MAIN}}}calcPr")
            if calc_pr is None:
                calc_pr = etree.SubElement(root, f"{{{NS_MAIN}}}calcPr")
            calc_pr.set("fullCalcOnLoad", "1")

        return etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )