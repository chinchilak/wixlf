import shutil
import tempfile
import zipfile
from collections.abc import Iterable, Mapping
from pathlib import Path
from typing import Literal

from lxml import etree

from .assembler import PackageAssembler
from .constants import MAX_EXCEL_COLS, MAX_EXCEL_ROWS, NS_MAIN
from .models import (
    SheetEditPlan,
    SheetVisibilityMutation,
    SheetVisibilityState,
    WorksheetTaskPayload,
)
from .patcher import patch_sheet_worker
from .reader import PackageReader
from .utils.coords import parse_col_ref
from .utils.xml_text import parse_xml
from .utils.zip_io import clone_zip_info, stream_zip_entry

type RowTargets = Iterable[int] | Literal["all"]
type ColumnTargets = Iterable[str | int] | Literal["all"]


def _normalize_targets(
    targets: Iterable[int] | None,
    axis: str,
    maximum: int,
) -> tuple[int, ...] | None:
    if targets is None:
        return None
    values = tuple(targets)
    if any(not isinstance(value, int) or isinstance(value, bool) for value in values):
        raise TypeError(f"{axis} targets must be integers.")
    if any(value < 1 or value > maximum for value in values):
        raise ValueError(f"{axis} targets must be between 1 and {maximum}.")
    return tuple(sorted(set(values)))


def _normalize_row_targets(targets: RowTargets, hidden: bool) -> tuple[int, ...] | None:
    if targets == "all":
        if hidden:
            raise ValueError("'all' is supported only when unhiding rows.")
        return None
    if isinstance(targets, str):
        raise TypeError("Row targets must be an iterable of row numbers or 'all'.")
    return _normalize_targets(targets, "Row", MAX_EXCEL_ROWS)


def _normalize_column_targets(
    targets: ColumnTargets,
    hidden: bool,
) -> tuple[int, ...] | None:
    if targets == "all":
        if hidden:
            raise ValueError("'all' is supported only when unhiding columns.")
        return None
    if isinstance(targets, str):
        raise TypeError("Column targets must be an iterable of column references or 'all'.")
    values = tuple(parse_col_ref(value) for value in targets)
    if any(value > MAX_EXCEL_COLS for value in values):
        raise ValueError(f"Column targets must not exceed {MAX_EXCEL_COLS}.")
    return tuple(sorted(set(values)))


def _rewrite_sheet_states(
    src_path: Path,
    dst_path: Path,
    states: Mapping[str, SheetVisibilityState],
    compresslevel: int,
) -> None:
    with (
        zipfile.ZipFile(src_path, "r") as zin,
        zipfile.ZipFile(
            dst_path,
            "w",
            compression=zipfile.ZIP_DEFLATED,
            compresslevel=compresslevel,
        ) as zout,
    ):
        for info in zin.infolist():
            if info.filename == "xl/workbook.xml":
                root = parse_xml(zin.read(info.filename))
                sheets = root.findall(f".//{{{NS_MAIN}}}sheet")
                sheet_nodes = {sheet.get("name"): sheet for sheet in sheets}
                for name, state in states.items():
                    sheet = sheet_nodes[name]
                    if state == "visible":
                        sheet.attrib.pop("state", None)
                    else:
                        sheet.set("state", state)

                visible_indices = [
                    index
                    for index, sheet in enumerate(sheets)
                    if sheet.get("state", "visible") == "visible"
                ]
                if not visible_indices:
                    raise ValueError("Workbook must retain at least one visible worksheet.")
                active_tab = visible_indices[0]
                for workbook_view in root.iter(f"{{{NS_MAIN}}}workbookView"):
                    raw_active = workbook_view.get("activeTab")
                    if raw_active is None or not raw_active.isdigit():
                        continue
                    index = int(raw_active)
                    if index >= len(sheets) or sheets[index].get("state", "visible") != "visible":
                        workbook_view.set("activeTab", str(active_tab))

                zout.writestr(
                    clone_zip_info(info),
                    etree.tostring(
                        root,
                        xml_declaration=True,
                        encoding="UTF-8",
                        standalone=True,
                    ),
                    compresslevel=compresslevel,
                )
                continue
            stream_zip_entry(zin, zout, info, 1024 * 1024, compresslevel)


def set_visibility(
    src_path: str | Path,
    *,
    sheets: Mapping[str, SheetVisibilityState] | None = None,
    rows: Mapping[str, RowTargets] | None = None,
    columns: Mapping[str, ColumnTargets] | None = None,
    hidden: bool = False,
    dst_path: str | Path | None = None,
    max_workers: int | None = None,
    compresslevel: int = 1,
) -> None:
    """Sets visibility for worksheet tabs, rows, or columns in an existing workbook."""
    src = Path(src_path)
    dst = Path(dst_path) if dst_path is not None else src
    sheet_infos = PackageReader(src).get_sheets()
    sheet_states = dict(sheets or {})
    row_targets = dict(rows or {})
    column_targets = dict(columns or {})

    for name in set(sheet_states) | set(row_targets) | set(column_targets):
        if name not in sheet_infos:
            raise KeyError(f"Worksheet {name!r} not found in {src}")
    for name, state in sheet_states.items():
        if state not in ("visible", "hidden", "veryHidden"):
            raise ValueError(f"Invalid worksheet visibility state: {state!r}")

    visible_count = sum(
        1
        for name, sheet in sheet_infos.items()
        if sheet_states.get(name, sheet.state) == "visible"
    )
    if sheet_states and visible_count == 0:
        raise ValueError("Workbook must retain at least one visible worksheet.")

    visibility_by_sheet: dict[str, SheetVisibilityMutation] = {}
    for name in set(row_targets) | set(column_targets):
        row_indices = (
            _normalize_row_targets(row_targets[name], hidden)
            if name in row_targets
            else ()
        )
        column_indices = (
            _normalize_column_targets(column_targets[name], hidden)
            if name in column_targets
            else ()
        )
        if row_indices or row_indices is None or column_indices or column_indices is None:
            visibility_by_sheet[name] = SheetVisibilityMutation(
                rows=row_indices,
                columns=column_indices,
                hidden=hidden,
            )

    if not sheet_states and not visibility_by_sheet:
        raise ValueError("Specify at least one worksheet, row, or column visibility change.")

    dst.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix=".wixlf_visibility_", dir=dst.parent) as temp_dir:
        staged = Path(temp_dir) / "workbook.xlsx"
        if visibility_by_sheet:
            dirty_edits = {
                sheet_infos[name].zip_path: WorksheetTaskPayload(
                    plan=SheetEditPlan.empty(),
                    visibility=mutation,
                )
                for name, mutation in visibility_by_sheet.items()
            }
            PackageAssembler(
                src_path=src,
                dst_path=staged,
                max_workers=max_workers,
                compresslevel=compresslevel,
            ).build(
                dirty_edits=dirty_edits,
                patch_fn=patch_sheet_worker,
                force_recalc=False,
            )
        else:
            shutil.copy2(src, staged)

        final_staged = staged
        if sheet_states:
            states_path = Path(temp_dir) / "workbook_states.xlsx"
            _rewrite_sheet_states(staged, states_path, sheet_states, compresslevel)
            final_staged = states_path

        shutil.copymode(src, final_staged)
        final_staged.replace(dst)


def unhide_all_sheets(
    src_path: str | Path,
    *,
    dst_path: str | Path | None = None,
    max_workers: int | None = None,
    compresslevel: int = 1,
) -> None:
    """Unhides every worksheet tab without changing row or column visibility."""
    sheet_names = PackageReader(src_path).get_sheets()
    sheet_states: dict[str, SheetVisibilityState] = {
        name: "visible" for name in sheet_names
    }
    set_visibility(
        src_path,
        sheets=sheet_states,
        dst_path=dst_path,
        max_workers=max_workers,
        compresslevel=compresslevel,
    )


def unhide_all_rows_and_columns(
    src_path: str | Path,
    *,
    dst_path: str | Path | None = None,
    max_workers: int | None = None,
    compresslevel: int = 1,
) -> None:
    """Unhides every row and column on every worksheet without changing sheet-tab state."""
    sheet_names = PackageReader(src_path).get_sheets()
    unhide_targets = {name: "all" for name in sheet_names}
    set_visibility(
        src_path,
        rows=unhide_targets,
        columns=unhide_targets,
        hidden=False,
        dst_path=dst_path,
        max_workers=max_workers,
        compresslevel=compresslevel,
    )
