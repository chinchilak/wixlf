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
)
from .patcher import SheetStreamPatcher, create_sheet_worker, patch_sheet_worker
from .reader import PackageReader

__version__ = "0.2.0"

__all__ = [
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
    "__version__",
    "add_sheet",
    "create_sheet_worker",
    "delete_sheets",
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
    max_workers: int | None = None,
    force_recalc: bool = True,
) -> None:
    """
    Modifies an existing workbook in-place by default (or writes to dst_path if given).
    Supports cell updates, creating new sheets, deleting sheets, and reordering tabs.
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
    )
    assembler.build(
        dirty_edits=dirty_edits,
        patch_fn=patch_sheet_worker,
        new_sheets=normalized_new,
        delete_sheets=delete_sheets,
        move_sheets=normalized_moves,
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
    )


def delete_sheets(
    src_path: str | Path,
    names: str | Iterable[str],
    *,
    dst_path: str | Path | None = None,
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
    )


def move_sheet(
    src_path: str | Path,
    name: str,
    position: int | None = None,
    *,
    dst_path: str | Path | None = None,
) -> None:
    """
    Moves an existing worksheet to 0-based `position` (or to the end if position=None).
    Modifies `src_path` in-place unless `dst_path` is specified.
    """
    patch_workbook(
        src_path=src_path,
        dst_path=dst_path,
        move_sheets=[SheetMoveSpec(name=name, position=position)],
    )
