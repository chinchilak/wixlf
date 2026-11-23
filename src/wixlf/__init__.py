from collections.abc import Iterable, Mapping, Sequence
from dataclasses import replace
from pathlib import Path

from .assembler import PackageAssembler
from .constants import CellScalar, PatchWorkerFn
from .models import (
    CellEdit,
    ClearRangeSpec,
    Formula,
    NewSheetSpec,
    RowEditBatch,
    SheetEditPlan,
    SheetInfo,
    SheetMoveSpec,
    SheetStructureMutation,
    SheetVisibilityMutation,
    SheetVisibilityState,
    WorksheetTaskPayload,
)
from .patcher import SheetStreamPatcher, create_sheet_worker, patch_sheet_worker
from .reader import PackageReader
from .utils.coords import check_cell_bounds, parse_col_ref, parse_range
from .utils.names import validate_sheet_name
from .utils.shifting import AxisShiftOp
from .visibility import set_visibility, unhide_all_rows_and_columns, unhide_all_sheets

__version__ = "0.4.0"

__all__ = [
    "AxisShiftOp",
    "CellEdit",
    "CellScalar",
    "ClearRangeSpec",
    "Formula",
    "NewSheetSpec",
    "PackageAssembler",
    "PackageReader",
    "PatchWorkerFn",
    "RowEditBatch",
    "SheetEditPlan",
    "SheetInfo",
    "SheetMoveSpec",
    "SheetStreamPatcher",
    "SheetStructureMutation",
    "SheetVisibilityMutation",
    "SheetVisibilityState",
    "WorksheetTaskPayload",
    "__version__",
    "add_sheet",
    "clear_cells",
    "create_sheet_worker",
    "delete_cols",
    "delete_rows",
    "delete_sheets",
    "insert_cols",
    "insert_rows",
    "move_sheet",
    "patch_sheet_worker",
    "patch_workbook",
    "set_visibility",
    "unhide_all_rows_and_columns",
    "unhide_all_sheets",
    "validate_sheet_name",
]

type CellMapInput = Mapping[str | tuple[int, int], CellScalar] | SheetEditPlan


def _to_plan(data: CellMapInput | None, formulas_from_strings: bool = False) -> SheetEditPlan:
    if data is None:
        return SheetEditPlan.empty()
    plan = data if isinstance(data, SheetEditPlan) else SheetEditPlan.from_dict(data)
    return _promote_string_formulas(plan) if formulas_from_strings else plan


def _promote_string_formulas(plan: SheetEditPlan) -> SheetEditPlan:
    """Legacy mode: converts str values starting with '=' into Formula values."""
    return replace(
        plan,
        rows=tuple(
            RowEditBatch(
                row=batch.row,
                cells=tuple(
                    replace(cell, value=Formula(cell.value))
                    if isinstance(cell.value, str) and cell.value.startswith("=")
                    else cell
                    for cell in batch.cells
                ),
            )
            for batch in plan.rows
        )
    )


def patch_workbook(
    src_path: str | Path,
    edits_by_sheet: Mapping[str, CellMapInput] | None = None,
    *,
    dst_path: str | Path | None = None,
    new_sheets: Mapping[str, CellMapInput] | Sequence[NewSheetSpec] | None = None,
    delete_sheets: Iterable[str] = (),
    move_sheets: Sequence[SheetMoveSpec | tuple[str, int | None]] = (),
    shift_ops: Mapping[str, tuple[Sequence[AxisShiftOp], Sequence[AxisShiftOp]]] | None = None,
    update_formulas: bool = True,
    max_workers: int | None = None,
    compresslevel: int = 1,
    force_recalc: bool = True,
    formulas_from_strings: bool = False,
) -> None:
    """
    Modifies an existing workbook in-place by default (or writes to dst_path if given).
    Supports cell updates, row/col shifting, table resizing, formula rewriting,
    creating new sheets, deleting sheets, and reordering tabs.
    Formulas must be passed as Formula(...) unless formulas_from_strings=True,
    in which case str values starting with '=' are also treated as formulas.
    Edits combined with shift_ops are expressed in post-shift coordinates
    (including formula references).
    """
    reader = PackageReader(src_path)
    sheets = reader.get_sheets()

    dirty_edits: dict[str, SheetEditPlan] = {}
    if edits_by_sheet:
        for sheet_name, edits in edits_by_sheet.items():
            if sheet_name not in sheets:
                raise KeyError(f"Worksheet {sheet_name!r} not found in {src_path}")
            dirty_edits[sheets[sheet_name].zip_path] = _to_plan(edits, formulas_from_strings)

    normalized_new: list[NewSheetSpec] = []
    if new_sheets:
        if isinstance(new_sheets, Mapping):
            for name, data in new_sheets.items():
                normalized_new.append(
                    NewSheetSpec(name=name, plan=_to_plan(data, formulas_from_strings))
                )
        else:
            normalized_new.extend(
                replace(spec, plan=_to_plan(spec.plan, formulas_from_strings))
                for spec in new_sheets
            )

    normalized_moves: list[SheetMoveSpec] = [
        item if isinstance(item, SheetMoveSpec) else SheetMoveSpec(name=item[0], position=item[1])
        for item in move_sheets
    ]

    assembler = PackageAssembler(
        src_path=src_path,
        dst_path=dst_path,
        max_workers=max_workers,
        compresslevel=compresslevel,
    )
    assembler.build(
        dirty_edits=dirty_edits,
        patch_fn=patch_sheet_worker,
        new_sheets=normalized_new,
        delete_sheets=delete_sheets,
        move_sheets=normalized_moves,
        shift_ops=shift_ops,
        update_formulas=update_formulas,
        force_recalc=force_recalc,
    )


def add_sheet(
    src_path: str | Path,
    name: str,
    data: CellMapInput | None = None,
    *,
    position: int | None = None,
    dst_path: str | Path | None = None,
    max_workers: int | None = None,
    compresslevel: int = 1,
    formulas_from_strings: bool = False,
) -> None:
    """
    Creates a new worksheet at 0-based `position` (or at the end if position=None).
    Modifies `src_path` in-place unless `dst_path` is specified.
    """
    spec = NewSheetSpec(
        name=name, plan=_to_plan(data, formulas_from_strings), position=position
    )
    patch_workbook(
        src_path=src_path,
        dst_path=dst_path,
        new_sheets=[spec],
        max_workers=max_workers,
        compresslevel=compresslevel,
    )


def delete_sheets(
    src_path: str | Path,
    names: str | Iterable[str],
    *,
    dst_path: str | Path | None = None,
    update_formulas: bool = True,
    compresslevel: int = 1,
) -> None:
    """
    Deletes one or more worksheets by name.
    Modifies `src_path` in-place unless `dst_path` is specified.
    """
    target_names = [names] if isinstance(names, str) else list(names)
    patch_workbook(
        src_path=src_path,
        dst_path=dst_path,
        delete_sheets=target_names,
        update_formulas=update_formulas,
        compresslevel=compresslevel,
    )


def move_sheet(
    src_path: str | Path,
    name: str,
    position: int | None = None,
    *,
    dst_path: str | Path | None = None,
    compresslevel: int = 1,
) -> None:
    """
    Moves an existing worksheet to 0-based `position` (or to the end if position=None).
    Modifies `src_path` in-place unless `dst_path` is specified.
    """
    patch_workbook(
        src_path=src_path,
        dst_path=dst_path,
        move_sheets=[SheetMoveSpec(name=name, position=position)],
        compresslevel=compresslevel,
    )


def insert_rows(
    src_path: str | Path,
    sheet: str,
    row_idx: int,
    count: int = 1,
    *,
    dst_path: str | Path | None = None,
    update_formulas: bool = True,
    max_workers: int | None = None,
    compresslevel: int = 1,
) -> None:
    """
    Inserts `count` blank rows starting at 1-based `row_idx`, shifting existing
    rows, formulas, merged cells, Excel Tables, and defined names down.
    """
    if row_idx < 1 or count < 1:
        raise ValueError("row_idx and count must be >= 1")
    patch_workbook(
        src_path=src_path,
        dst_path=dst_path,
        shift_ops={sheet: ([AxisShiftOp(start=row_idx, delta=count)], [])},
        update_formulas=update_formulas,
        max_workers=max_workers,
        compresslevel=compresslevel,
    )


def delete_rows(
    src_path: str | Path,
    sheet: str,
    row_idx: int,
    count: int = 1,
    *,
    dst_path: str | Path | None = None,
    update_formulas: bool = True,
    max_workers: int | None = None,
    compresslevel: int = 1,
) -> None:
    """
    Deletes `count` rows starting at 1-based `row_idx`, shifting downstream
    rows, formulas (#REF! if targeting deleted rows), tables, and ranges up.
    """
    if row_idx < 1 or count < 1:
        raise ValueError("row_idx and count must be >= 1")
    patch_workbook(
        src_path=src_path,
        dst_path=dst_path,
        shift_ops={sheet: ([AxisShiftOp(start=row_idx, delta=-count)], [])},
        update_formulas=update_formulas,
        max_workers=max_workers,
        compresslevel=compresslevel,
    )


def insert_cols(
    src_path: str | Path,
    sheet: str,
    col: str | int,
    count: int = 1,
    *,
    dst_path: str | Path | None = None,
    update_formulas: bool = True,
    max_workers: int | None = None,
    compresslevel: int = 1,
) -> None:
    """
    Inserts `count` blank columns starting at `col` (1-based int or 'B' letters),
    shifting existing columns, widths, formulas, tables, and ranges right.
    """
    if count < 1:
        raise ValueError("count must be >= 1")
    col_idx = parse_col_ref(col)
    patch_workbook(
        src_path=src_path,
        dst_path=dst_path,
        shift_ops={sheet: ([], [AxisShiftOp(start=col_idx, delta=count)])},
        update_formulas=update_formulas,
        max_workers=max_workers,
        compresslevel=compresslevel,
    )


def delete_cols(
    src_path: str | Path,
    sheet: str,
    col: str | int,
    count: int = 1,
    *,
    dst_path: str | Path | None = None,
    update_formulas: bool = True,
    max_workers: int | None = None,
    compresslevel: int = 1,
) -> None:
    """
    Deletes `count` columns starting at `col` (1-based int or 'B' letters),
    shifting remaining columns, widths, formulas, tables, and ranges left.
    """
    if count < 1:
        raise ValueError("count must be >= 1")
    col_idx = parse_col_ref(col)
    patch_workbook(
        src_path=src_path,
        dst_path=dst_path,
        shift_ops={sheet: ([], [AxisShiftOp(start=col_idx, delta=-count)])},
        update_formulas=update_formulas,
        max_workers=max_workers,
        compresslevel=compresslevel,
    )


def clear_cells(
    src_path: str | Path,
    sheet: str,
    targets: str | Iterable[str | tuple[int, int]],
    *,
    clear_style: bool = False,
    drop_cells: bool = False,
    dst_path: str | Path | None = None,
    max_workers: int | None = None,
    compresslevel: int = 1,
) -> None:
    """
    Clears values in one or more cells or A1:D10 ranges without shifting rows/cols.
    Only cells that already exist are touched; memory use is independent of range size.
    - clear_style=True also strips cell formatting ('s').
    - drop_cells=True removes the <c> XML tags completely.
    """
    items = [targets] if isinstance(targets, str) else list(targets)
    ranges: list[ClearRangeSpec] = []

    for item in items:
        if isinstance(item, tuple):
            r, c = item
            check_cell_bounds(r, c)
            r1, c1, r2, c2 = r, c, r, c
        else:
            r1, c1, r2, c2 = parse_range(item)
        ranges.append(
            ClearRangeSpec(
                min_row=r1,
                min_col=c1,
                max_row=r2,
                max_col=c2,
                clear_style=clear_style,
                drop_cell=drop_cells,
            )
        )

    patch_workbook(
        src_path=src_path,
        edits_by_sheet={sheet: SheetEditPlan(rows=(), clear_ranges=tuple(ranges))},
        dst_path=dst_path,
        max_workers=max_workers,
        compresslevel=compresslevel,
    )
