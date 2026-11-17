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
    CALC_CHAIN_CT,
    CALC_CHAIN_PATH,
    CALC_CHAIN_REL_TYPE,
    NS_CT,
    NS_MAIN,
    NS_REL_DOC,
    NS_REL_PKG,
    TABLE_REL_TYPE,
    TAG_AUTO_FILTER,
    TAG_DEFINED_NAME,
    TAG_DEFINED_NAMES,
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
from .utils.shifting import AxisShiftOp, rewrite_formula, shift_range_str
from .utils.zip_io import (
    clone_zip_info,
    copy_disk_file_to_zip,
    new_zip_info,
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
    fd, tmp_xml_path = tempfile.mkstemp(
        prefix="sheet_patch_", suffix=".xml", dir=temp_dir
    )
    os.close(fd)
    try:
        patch_fn(src_path, sheet_zip_path, edits, tmp_xml_path)
        return sheet_zip_path, tmp_xml_path
    except Exception:
        if os.path.exists(tmp_xml_path):
            os.remove(tmp_xml_path)
        raise


class PackageAssembler:
    """
    Coordinates multiprocessing across dirty and newly created sheets, performs
    sheet structural operations (create, delete, reorder, row/col shifting,
    Excel Table resizing, and formula rewriting), and atomically writes the
    resulting .xlsx/.xlsm archive (in-place by default, or to dst_path).
    """

    # Avoid process spawn overhead when modifying a single tiny worksheet (< 512 KB)
    PARALLEL_SIZE_THRESHOLD_BYTES = 512 * 1024

    def __init__(
        self,
        src_path: str | Path,
        dst_path: str | Path | None = None,
        max_workers: int | None = None,
        chunk_size: int = 1024 * 1024,  # 1 MB streaming buffer
        compresslevel: int = 1,
    ) -> None:
        self.src_path = Path(src_path)
        self.dst_path = Path(dst_path) if dst_path is not None else self.src_path
        self.max_workers = max_workers or os.process_cpu_count() or 1
        self.chunk_size = chunk_size
        self.compresslevel = compresslevel

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
                deleted_table_zips,
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
                deleted_table_zips=deleted_table_zips,
                table_shifts_by_zip=table_shifts_by_zip,
                delete_sheet_names=delete_set,
                move_sheets=move_sheets,
                shift_ops=normalized_shifts if update_formulas else {},
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
            wb_rels_root = etree.fromstring(zin.read("xl/_rels/workbook.xml.rels"))
            wb_root = etree.fromstring(zin.read("xl/workbook.xml"))

            rel_map: dict[str, str] = {
                rel_id: target
                for rel in wb_rels_root.findall(f"{{{NS_REL_PKG}}}Relationship")
                if (rel_id := rel.get("Id")) and (target := rel.get("Target"))
            }

            existing_sheets_by_name: dict[str, tuple[int, str]] = {}
            max_sheet_id = 0
            for sheet_elem in wb_root.findall(f".//{{{NS_MAIN}}}sheet"):
                name = sheet_elem.get("name", "")
                sid_str = sheet_elem.get("sheetId", "0")
                rid = sheet_elem.get(f"{{{NS_REL_DOC}}}id", "")
                if sid_str.isdigit():
                    max_sheet_id = max(max_sheet_id, int(sid_str))
                if name and rid in rel_map:
                    zip_path = resolve_zip_path("xl/workbook.xml", rel_map[rid])
                    existing_sheets_by_name[name] = (int(sid_str or 0), zip_path)

            # Validate deletions
            unknown_deletes = delete_set - existing_sheets_by_name.keys()
            if unknown_deletes:
                raise KeyError(
                    f"Cannot delete unknown worksheet(s): {sorted(unknown_deletes)}"
                )

            remaining_count = (
                len(existing_sheets_by_name) - len(delete_set) + len(new_sheets)
            )
            if remaining_count < 1:
                raise ValueError("Workbook must retain at least one visible worksheet.")

            deleted_zip_paths = {
                existing_sheets_by_name[name][1] for name in delete_set
            }
            surviving_sheets = {
                name: info[1]
                for name, info in existing_sheets_by_name.items()
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
            deleted_table_zips: set[str] = set()

            for s_name, (_, s_zip_path) in existing_sheets_by_name.items():
                s_rels_path = self._rels_path_for_sheet(s_zip_path)
                if s_rels_path not in existing_entries:
                    continue
                s_rels_root = etree.fromstring(zin.read(s_rels_path))
                for rel in s_rels_root.findall(f"{{{NS_REL_PKG}}}Relationship"):
                    if rel.get("Type") == TABLE_REL_TYPE and (t := rel.get("Target")):
                        table_zip = resolve_zip_path(s_zip_path, t)
                        if s_name in delete_set:
                            deleted_table_zips.add(table_zip)
                        elif s_name in shift_ops:
                            table_shifts_by_zip[table_zip] = shift_ops[s_name]

        # Validate new sheet names against surviving + newly added names
        surviving_names = set(surviving_sheets.keys())
        seen_new_names: set[str] = set()
        for spec in new_sheets:
            if not spec.name:
                raise ValueError("New worksheet name cannot be empty.")
            if spec.name in surviving_names or spec.name in seen_new_names:
                raise ValueError(f"Duplicate worksheet name: {spec.name!r}")
            seen_new_names.add(spec.name)

        # Validate move targets
        final_names = surviving_names | seen_new_names
        for move in move_sheets:
            if move.name not in final_names:
                raise KeyError(f"Cannot move unknown worksheet: {move.name!r}")

        # Allocate unique sheetN.xml, sheetId, and rId for each new sheet
        used_rids = set(rel_map.keys())
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
            deleted_table_zips,
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
            dirty_sizes = [
                zin.getinfo(path).file_size for path in dirty_edits
            ]

        use_multiprocessing = len(tasks) > 1 or any(
            sz >= self.PARALLEL_SIZE_THRESHOLD_BYTES for sz in dirty_sizes
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
        deleted_table_zips: set[str],
        table_shifts_by_zip: Mapping[
            str, tuple[tuple[AxisShiftOp, ...], tuple[AxisShiftOp, ...]]
        ],
        delete_sheet_names: set[str],
        move_sheets: Sequence[SheetMoveSpec],
        shift_ops: Mapping[str, tuple[tuple[AxisShiftOp, ...], tuple[AxisShiftOp, ...]]],
        strip_calc_chain: bool,
    ) -> None:
        """
        Streams entries into the destination archive without buffering whole files,
        using fast deflate compression (compresslevel) across all streamed parts.
        """
        dst_dir = self.dst_path.parent
        dst_dir.mkdir(parents=True, exist_ok=True)
        fd, staging_zip = tempfile.mkstemp(
            prefix=".tmp_xlsx_", suffix=".xlsx", dir=dst_dir
        )
        os.close(fd)

        deleted_rels_paths = {
            self._rels_path_for_sheet(p) for p in deleted_zip_paths
        }
        all_deleted_parts = deleted_zip_paths | deleted_table_zips

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
                for info in zin.infolist():
                    name = info.filename

                    if strip_calc_chain and name == CALC_CHAIN_PATH:
                        continue

                    if name in all_deleted_parts or name in deleted_rels_paths:
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

            os.replace(staging_zip, self.dst_path)
        except Exception:
            if os.path.exists(staging_zip):
                os.remove(staging_zip)
            raise

    @staticmethod
    def _rels_path_for_sheet(sheet_zip_path: str) -> str:
        """Converts 'xl/worksheets/sheet1.xml' -> 'xl/worksheets/_rels/sheet1.xml.rels'."""
        prefix, _, filename = sheet_zip_path.rpartition("/")
        return f"{prefix}/_rels/{filename}.rels" if prefix else f"_rels/{filename}.rels"

    @staticmethod
    def _update_table_xml(
        xml_bytes: bytes,
        row_ops: Sequence[AxisShiftOp],
        col_ops: Sequence[AxisShiftOp],
    ) -> bytes:
        """Resizes <table ref="..."> and nested <autoFilter ref="..."> on tableN.xml."""
        root = etree.fromstring(xml_bytes)
        ref = root.get("ref", "")
        if ref:
            new_ref = shift_range_str(ref, row_ops, col_ops)
            if new_ref is not None:
                root.set("ref", new_ref)

        af = root.find(TAG_AUTO_FILTER)
        if af is not None:
            af_ref = af.get("ref", "")
            if af_ref:
                new_af = shift_range_str(af_ref, row_ops, col_ops)
                if new_af is not None:
                    af.set("ref", new_af)

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
        root = etree.fromstring(xml_bytes)
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
        root = etree.fromstring(xml_bytes)

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
    def _update_workbook_xml(
        xml_bytes: bytes,
        allocated_new: Sequence[_AllocatedNewSheet],
        delete_sheet_names: set[str],
        move_sheets: Sequence[SheetMoveSpec],
        shift_ops: Mapping[str, tuple[tuple[AxisShiftOp, ...], tuple[AxisShiftOp, ...]]],
        force_recalc: bool,
    ) -> bytes:
        root = etree.fromstring(xml_bytes)
        sheets_elem = root.find(f"{{{NS_MAIN}}}sheets")
        if sheets_elem is None:
            sheets_elem = etree.SubElement(root, f"{{{NS_MAIN}}}sheets")

        # 1. Remove deleted <sheet> entries
        sheet_nodes: list[etree._Element] = []
        for s_elem in list(sheets_elem.findall(f"{{{NS_MAIN}}}sheet")):
            if s_elem.get("name") not in delete_sheet_names:
                sheet_nodes.append(s_elem)

        # 2. Insert newly created <sheet> elements at requested positions (or end)
        for alloc in allocated_new:
            new_elem = etree.Element(
                f"{{{NS_MAIN}}}sheet",
                attrib={
                    "name": alloc.name,
                    "sheetId": alloc.sheet_id,
                    f"{{{NS_REL_DOC}}}id": alloc.rel_id,
                },
            )
            if alloc.position is None:
                sheet_nodes.append(new_elem)
            else:
                clamped_pos = max(0, min(alloc.position, len(sheet_nodes)))
                sheet_nodes.insert(clamped_pos, new_elem)

        # 3. Apply sequential sheet tab moves
        for move in move_sheets:
            idx = next(
                (i for i, el in enumerate(sheet_nodes) if el.get("name") == move.name),
                None,
            )
            if idx is None:
                continue
            node = sheet_nodes.pop(idx)
            if move.position is None:
                sheet_nodes.append(node)
            else:
                clamped_pos = max(0, min(move.position, len(sheet_nodes)))
                sheet_nodes.insert(clamped_pos, node)

        # Replace <sheets> children in final tab order
        sheets_elem[:] = sheet_nodes

        # 4. Update <definedNames> formulas if any rows/cols shifted or sheets deleted
        if shift_ops or delete_sheet_names:
            defined_names = root.find(TAG_DEFINED_NAMES)
            if defined_names is not None:
                for dn in list(defined_names.findall(TAG_DEFINED_NAME)):
                    if dn.text:
                        local_idx = dn.get("localSheetId")
                        cur_sheet = ""
                        if local_idx is not None and local_idx.isdigit():
                            idx_val = int(local_idx)
                            if 0 <= idx_val < len(sheet_nodes):
                                cur_sheet = sheet_nodes[idx_val].get("name", "")
                        dn.text = rewrite_formula(
                            dn.text,
                            current_sheet=cur_sheet,
                            shifts_by_sheet=shift_ops,
                            deleted_sheets=delete_sheet_names,
                        )

        # Reset activeTab in <bookViews>/<workbookView> so it stays within bounds
        wb_view = root.find(f".//{{{NS_MAIN}}}workbookView")
        if wb_view is not None and "activeTab" in wb_view.attrib:
            active_idx = int(wb_view.get("activeTab", "0") or 0)
            if active_idx >= len(sheet_nodes):
                wb_view.set("activeTab", "0")

        if force_recalc:
            calc_pr = root.find(f"{{{NS_MAIN}}}calcPr")
            if calc_pr is None:
                calc_pr = etree.SubElement(root, f"{{{NS_MAIN}}}calcPr")
            calc_pr.set("fullCalcOnLoad", "1")

        return etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )