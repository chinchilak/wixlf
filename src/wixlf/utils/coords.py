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


def parse_coord(coord: str) -> tuple[int, int]:
    """
    Parses 'AA105' or '$AA$105' into (row_idx, col_idx) 1-based integers
    using pure string/character operations without regex.
    """
    cleaned = coord.strip()
    if cleaned.startswith("$"):
        cleaned = cleaned[1:]

    split_idx = 0
    length = len(cleaned)
    while split_idx < length and cleaned[split_idx].isascii() and cleaned[split_idx].isalpha():
        split_idx += 1

    col_letters = cleaned[:split_idx]
    row_part = cleaned[split_idx:]
    if row_part.startswith("$"):
        row_part = row_part[1:]

    if (
        not (1 <= len(col_letters) <= 3)
        or not row_part
        or not (row_part.isascii() and row_part.isdigit())
    ):
        raise ValueError(f"Invalid Excel cell coordinate: {coord!r}")

    row_idx = int(row_part)
    if row_idx < 1:
        raise ValueError(f"Invalid Excel cell coordinate: {coord!r}")

    return row_idx, col_to_idx(col_letters)


def format_coord(row_idx: int, col_idx: int) -> str:
    """Formats 1-based (row_idx, col_idx) into 'AA105'."""
    return f"{idx_to_col(col_idx)}{row_idx}"
