from collections.abc import Mapping
from dataclasses import dataclass

from .constants import CellScalar
from .utils.coords import parse_coord


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
