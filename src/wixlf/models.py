from collections.abc import Mapping, Sequence
from dataclasses import dataclass

from .constants import CellScalar
from .utils.coords import parse_coord
from .utils.shifting import AxisShiftOp


@dataclass(frozen=True, slots=True)
class SheetInfo:
    """Metadata mapping a human-readable worksheet name to its internal zip entry."""

    name: str
    sheet_id: str
    rel_id: str
    zip_path: str  # e.g., "xl/worksheets/sheet1.xml"


@dataclass(frozen=True, slots=True)
class CellEdit:
    """Represents a single cell mutation."""

    row: int
    col: int
    value: CellScalar
    style_id: str | None = None  # Optional explicit 's' attribute override
    clear_style: bool = False
    drop_cell: bool = False  # If True, omits <c> completely from the XML


@dataclass(frozen=True, slots=True)
class RowEditBatch:
    """Sorted cell mutations belonging to a single row."""

    row: int
    cells: tuple[CellEdit, ...]  # Strictly sorted by col ascending


@dataclass(frozen=True, slots=True)
class SheetEditPlan:
    """
    Pre-sorted, immutable, picklable edit plan for a single worksheet.
    Passed safely across ProcessPoolExecutor boundaries.
    """

    rows: tuple[RowEditBatch, ...]  # Strictly sorted by row ascending

    @classmethod
    def empty(cls) -> "SheetEditPlan":
        """Returns an empty edit plan (used when creating a blank worksheet)."""
        return cls(rows=())

    @classmethod
    def from_dict(
        cls,
        cell_map: Mapping[str | tuple[int, int], CellScalar],
        styles: Mapping[str | tuple[int, int], str] | None = None,
    ) -> "SheetEditPlan":
        """
        Compiles an unsorted mapping like {'B5': 42, (1, 1): 'Title'} into a
        strictly sorted SheetEditPlan ready for single-pass streaming.
        """
        by_row: dict[int, dict[int, CellEdit]] = {}
        style_lookup: dict[tuple[int, int], str] = {}

        if styles:
            for key, s_id in styles.items():
                r, c = parse_coord(key) if isinstance(key, str) else key
                style_lookup[(r, c)] = str(s_id)

        for key, val in cell_map.items():
            r, c = parse_coord(key) if isinstance(key, str) else key
            if r < 1 or c < 1:
                raise ValueError(f"Row and column must be >= 1, got ({r}, {c})")
            by_row.setdefault(r, {})[c] = CellEdit(
                row=r,
                col=c,
                value=val,
                style_id=style_lookup.get((r, c)),
            )

        sorted_batches = tuple(
            RowEditBatch(
                row=r,
                cells=tuple(by_row[r][c] for c in sorted(by_row[r])),
            )
            for r in sorted(by_row)
        )
        return cls(rows=sorted_batches)


@dataclass(frozen=True, slots=True)
class NewSheetSpec:
    """
    Specification for creating a brand-new worksheet in the workbook.
    position: 0-based tab index; None appends to the end.
    """

    name: str
    plan: SheetEditPlan
    position: int | None = None


@dataclass(frozen=True, slots=True)
class SheetMoveSpec:
    """
    Specification for repositioning an existing or newly created worksheet tab.
    position: 0-based target index; None moves the sheet to the end.
    """

    name: str
    position: int | None = None


@dataclass(frozen=True, slots=True)
class SheetStructureMutation:
    """
    Describes row/column insertions and deletions on a specific worksheet, plus
    workbook-wide formula shift context so cross-sheet formulas update cleanly.
    """

    sheet_name: str
    row_ops: tuple[AxisShiftOp, ...] = ()
    col_ops: tuple[AxisShiftOp, ...] = ()
    all_shifts: Mapping[str, tuple[Sequence[AxisShiftOp], Sequence[AxisShiftOp]]] | None = None
    deleted_sheets: tuple[str, ...] = ()

    @property
    def has_local_shifts(self) -> bool:
        return bool(self.row_ops or self.col_ops)

    @property
    def has_any_shifts(self) -> bool:
        if self.has_local_shifts or self.deleted_sheets:
            return True
        if self.all_shifts:
            return any(bool(r or c) for r, c in self.all_shifts.values())
        return False


@dataclass(frozen=True, slots=True)
class WorksheetTaskPayload:
    """Picklable payload passed to sheet worker processes."""

    plan: SheetEditPlan
    mutation: SheetStructureMutation | None = None
