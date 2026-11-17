from collections.abc import Iterable, Mapping, Sequence
from pathlib import Path

from .assembler import PackageAssembler
from .constants import CellScalar, PatchWorkerFn
from .models import (
    CellEdit,
    NewSheetSpec,
    RowEditBatch,
    SheetEditPlan,
    SheetInfo,
    SheetMoveSpec,
    SheetStructureMutation,
    WorksheetTaskPayload,
)
from .patcher import SheetStreamPatcher, create_sheet_worker, patch_sheet_worker
from .reader import PackageReader
from .utils.coords import parse_col_ref, parse_range
from .utils.shifting import AxisShiftOp

__version__ = "0.3.0"

__all__ = [
    "AxisShiftOp",
    "CellEdit",
    "CellScalar",
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
]

type CellMapInput = Mapping[str | tuple[int, int], CellScalar] | SheetEditPlan


def _to_plan(data: CellMapInput | None) -> SheetEditPlan:
    if data is None:
        return SheetEditPlan.empty()
    if isinstance(data, SheetEditPlan):
        return data
    return SheetEditPlan.from_dict(data)


def patch_workbook(
    src_path: str | Path,
    edits_by_sheet: Mapping[str, CellMapInput] | None = None,
    *,
    dst_path: str | Path | None = None,
    new_sheets: Mapping[str, CellMapInput] | Sequence[NewSheetSpec] | None = None,
    delete_sheets: Iterable[str] = (),
    move_sheets: Sequence[SheetMoveSpec | tuple[str, int | None]] = (),
    shift_ops: Mapping[str, tuple[Sequence[AxisShiftOp], Sequence[AxisShiftOp]]] | None = None,
    max_workers: int | None = None,
    compresslevel: int = 1,
    force_recalc: bool = True,
) -> None:
    """
    Modifies an existing workbook in-place by default (or writes to dst_path if given).
    Supports cell updates, row/col shifting, table resizing, formula rewriting,
    creating new sheets, deleting sheets, and reordering tabs.
    """
    reader = PackageReader(src_path)
    sheets = reader.get_sheets()

    dirty_edits: dict[str, SheetEditPlan] = {}
    if edits_by_sheet:
        for sheet_name, edits in edits_by_sheet.items():
            if sheet_name not in sheets:
                raise KeyError(f"Worksheet {sheet_name!r} not found in {src_path}")
            dirty_edits[sheets[sheet_name].zip_path] = _to_plan(edits)

    normalized_new: list[NewSheetSpec] = []
    if new_sheets:
        if isinstance(new_sheets, Mapping):
            for name, data in new_sheets.items():
                normalized_new.append(NewSheetSpec(name=name, plan=_to_plan(data)))
        else:
            normalized_new.extend(new_sheets)

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
) -> None:
    """
    Creates a new worksheet at 0-based `position` (or at the end if position=None).
    Modifies `src_path` in-place unless `dst_path` is specified.
    """
    spec = NewSheetSpec(name=name, plan=_to_plan(data), position=position)
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
    - clear_style=True also strips cell formatting ('s').
    - drop_cells=True removes the <c> XML tags completely.
    """
    items = [targets] if isinstance(targets, str) else list(targets)
    by_row: dict[int, dict[int, CellEdit]] = {}

    for item in items:
        if isinstance(item, tuple):
            r, c = item
            by_row.setdefault(r, {})[c] = CellEdit(
                row=r,
                col=c,
                value=None,
                clear_style=clear_style,
                drop_cell=drop_cells,
            )
        else:
            r1, c1, r2, c2 = parse_range(item)
            for r in range(r1, r2 + 1):
                row_bucket = by_row.setdefault(r, {})
                for c in range(c1, c2 + 1):
                    row_bucket[c] = CellEdit(
                        row=r,
                        col=c,
                        value=None,
                        clear_style=clear_style,
                        drop_cell=drop_cells,
                    )

    batches = tuple(
        RowEditBatch(row=r, cells=tuple(by_row[r][c] for c in sorted(by_row[r])))
        for r in sorted(by_row)
    )
    patch_workbook(
        src_path=src_path,
        edits_by_sheet={sheet: SheetEditPlan(rows=batches)},
        dst_path=dst_path,
        max_workers=max_workers,
        compresslevel=compresslevel,
    )
