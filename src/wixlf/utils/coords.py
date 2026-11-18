from ..constants import MAX_EXCEL_COLS, MAX_EXCEL_ROWS


def check_cell_bounds(row_idx: int, col_idx: int) -> None:
    """Raises ValueError if (row_idx, col_idx) lies outside the Excel grid."""
    if not (1 <= row_idx <= MAX_EXCEL_ROWS and 1 <= col_idx <= MAX_EXCEL_COLS):
        raise ValueError(
            f"Cell ({row_idx}, {col_idx}) is outside the Excel grid "
            f"(1..{MAX_EXCEL_ROWS} rows, 1..{MAX_EXCEL_COLS} cols)."
        )


def col_to_idx(col_str: str) -> int:
    """Converts column letters ('A' -> 1, 'Z' -> 26, 'AA' -> 27) to 1-based index."""
    idx = 0
    for ch in col_str.upper():
        idx = idx * 26 + (ord(ch) - 64)
    return idx


def idx_to_col(col_idx: int) -> str:
    """Converts 1-based column index (1 -> 'A', 27 -> 'AA') to column letters."""
    chars: list[str] = []
    n = col_idx
    while n > 0:
        n, rem = divmod(n - 1, 26)
        chars.append(chr(65 + rem))
    chars.reverse()
    return "".join(chars)


def parse_col_ref(col: str | int) -> int:
    """Normalizes a column index (1-based int) or column letters ('B', '$AA') to 1-based int."""
    if isinstance(col, int):
        if col < 1:
            raise ValueError(f"Column index must be >= 1, got {col}")
        return col
    cleaned = col.strip()
    cleaned = cleaned.removeprefix("$")
    if not (1 <= len(cleaned) <= 3) or not (cleaned.isascii() and cleaned.isalpha()):
        raise ValueError(f"Invalid Excel column reference: {col!r}")
    return col_to_idx(cleaned)


def parse_coord_anchored(coord: str) -> tuple[int, int, bool, bool]:
    """
    Parses 'AA105' or '$AA$105' into (row_idx, col_idx, col_abs, row_abs)
    using pure character operations without regex.
    """
    cleaned = coord.strip()
    col_abs = cleaned.startswith("$")
    if col_abs:
        cleaned = cleaned[1:]

    split_idx = 0
    length = len(cleaned)
    while (
        split_idx < length
        and cleaned[split_idx].isascii()
        and cleaned[split_idx].isalpha()
    ):
        split_idx += 1

    col_letters = cleaned[:split_idx]
    row_part = cleaned[split_idx:]
    row_abs = row_part.startswith("$")
    if row_abs:
        row_part = row_part[1:]

    if (
        not (1 <= len(col_letters) <= 3)
        or not row_part
        or not (row_part.isascii() and row_part.isdigit())
    ):
        raise ValueError(f"Invalid Excel cell coordinate: {coord!r}")

    row_idx = int(row_part)
    col_idx = col_to_idx(col_letters)
    if not (1 <= row_idx <= MAX_EXCEL_ROWS and col_idx <= MAX_EXCEL_COLS):
        raise ValueError(f"Invalid Excel cell coordinate: {coord!r}")

    return row_idx, col_idx, col_abs, row_abs


def parse_coord(coord: str) -> tuple[int, int]:
    """Parses 'AA105' or '$AA$105' into (row_idx, col_idx) 1-based integers."""
    row_idx, col_idx, _, _ = parse_coord_anchored(coord)
    return row_idx, col_idx


def format_coord(
    row_idx: int,
    col_idx: int,
    *,
    col_abs: bool = False,
    row_abs: bool = False,
) -> str:
    """Formats 1-based (row_idx, col_idx) into 'AA105' or '$AA$105'."""
    c_prefix = "$" if col_abs else ""
    r_prefix = "$" if row_abs else ""
    return f"{c_prefix}{idx_to_col(col_idx)}{r_prefix}{row_idx}"


def parse_range(range_ref: str) -> tuple[int, int, int, int]:
    """
    Parses 'A1:D10' or single cell 'B5' into (min_row, min_col, max_row, max_col).
    """
    part1, sep, part2 = range_ref.strip().partition(":")
    r1, c1 = parse_coord(part1)
    if not sep:
        return r1, c1, r1, c1
    r2, c2 = parse_coord(part2)
    return min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)


def format_range(
    min_row: int,
    min_col: int,
    max_row: int,
    max_col: int,
    *,
    c1_abs: bool = False,
    r1_abs: bool = False,
    c2_abs: bool = False,
    r2_abs: bool = False,
) -> str:
    """Formats bounding coordinates back to 'A1:D10' (or single cell if equal)."""
    start = format_coord(min_row, min_col, col_abs=c1_abs, row_abs=r1_abs)
    if min_row == max_row and min_col == max_col:
        return start
    end = format_coord(max_row, max_col, col_abs=c2_abs, row_abs=r2_abs)
    return f"{start}:{end}"
