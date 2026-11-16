from collections.abc import Mapping
from pathlib import Path

from .assembler import PackageAssembler
from .constants import CellScalar, PatchWorkerFn
from .models import CellEdit, RowEditBatch, SheetEditPlan, SheetInfo
from .patcher import SheetStreamPatcher, patch_sheet_worker
from .reader import PackageReader

__version__ = "0.1.0"

__all__ = [
    "CellEdit",
    "CellScalar",
    "PackageAssembler",
    "PackageReader",
    "PatchWorkerFn",
    "RowEditBatch",
    "SheetEditPlan",
    "SheetInfo",
    "SheetStreamPatcher",
    "__version__",
    "patch_sheet_worker",
    "patch_workbook",
]


def patch_workbook(
    src_path: str | Path,
    dst_path: str | Path,
    edits_by_sheet: Mapping[str, Mapping[str | tuple[int, int], CellScalar] | SheetEditPlan],
    *,
    max_workers: int | None = None,
    force_recalc: bool = True,
) -> None:
    """High-level convenience wrapper to patch multiple sheets by name."""
    reader = PackageReader(src_path)
    sheets = reader.get_sheets()

    dirty_edits: dict[str, SheetEditPlan] = {}
    for sheet_name, edits in edits_by_sheet.items():
        if sheet_name not in sheets:
            raise KeyError(f"Worksheet {sheet_name!r} not found in {src_path}")
        plan = (
            edits
            if isinstance(edits, SheetEditPlan)
            else SheetEditPlan.from_dict(edits)
        )
        dirty_edits[sheets[sheet_name].zip_path] = plan

    assembler = PackageAssembler(src_path, dst_path, max_workers=max_workers)
    assembler.build(
        dirty_edits=dirty_edits,
        patch_fn=patch_sheet_worker,
        force_recalc=force_recalc,
    )
