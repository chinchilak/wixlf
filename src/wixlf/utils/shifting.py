from collections.abc import Iterable, Mapping, Sequence
from dataclasses import dataclass

from ..constants import MAX_EXCEL_COLS, MAX_EXCEL_ROWS
from .coords import (
    col_to_idx,
    format_coord,
    format_range,
    idx_to_col,
    parse_coord_anchored,
)


@dataclass(frozen=True, slots=True)
class AxisShiftOp:
    """Represents a single insertion (delta > 0) or deletion (delta < 0) along rows or cols."""

    start: int  # 1-based row or column index where insertion/deletion begins
    delta: int  # > 0 for insert count, < 0 for delete count


def shift_index(idx: int, ops: Sequence[AxisShiftOp]) -> int | None:
    """
    Applies a sequence of AxisShiftOps to a 1-based coordinate index.
    Returns None if the coordinate falls inside a deleted span.
    """
    cur = idx
    for op in ops:
        if op.delta > 0:
            if cur >= op.start:
                cur += op.delta
        elif op.delta < 0:
            del_end = op.start + (-op.delta) - 1
            if op.start <= cur <= del_end:
                return None
            if cur > del_end:
                cur += op.delta
    return cur


def shift_range_bounds(
    min_r: int,
    min_c: int,
    max_r: int,
    max_c: int,
    row_ops: Sequence[AxisShiftOp],
    col_ops: Sequence[AxisShiftOp],
) -> tuple[int, int, int, int] | None:
    """
    Shifts a rectangular range (min_r, min_c, max_r, max_c) through row_ops and col_ops.
    Returns None if the entire range was deleted.
    """
    r1, r2 = min_r, max_r
    for op in row_ops:
        if op.delta > 0:
            if r1 >= op.start:
                r1 += op.delta
            if r2 >= op.start:
                r2 += op.delta
        elif op.delta < 0:
            del_start = op.start
            del_end = op.start + (-op.delta) - 1
            if r1 >= del_start and r2 <= del_end:
                return None
            # Count how many deleted rows lie strictly before r1 and up to r2
            before_r1 = max(0, min(r1 - 1, del_end) - del_start + 1)
            up_to_r2 = max(0, min(r2, del_end) - del_start + 1)
            r1 -= before_r1
            r2 -= up_to_r2
            if r1 > r2:
                return None

    c1, c2 = min_c, max_c
    for op in col_ops:
        if op.delta > 0:
            if c1 >= op.start:
                c1 += op.delta
            if c2 >= op.start:
                c2 += op.delta
        elif op.delta < 0:
            del_start = op.start
            del_end = op.start + (-op.delta) - 1
            if c1 >= del_start and c2 <= del_end:
                return None
            before_c1 = max(0, min(c1 - 1, del_end) - del_start + 1)
            up_to_c2 = max(0, min(c2, del_end) - del_start + 1)
            c1 -= before_c1
            c2 -= up_to_c2
            if c1 > c2:
                return None

    return (
        max(1, min(r1, MAX_EXCEL_ROWS)),
        max(1, min(c1, MAX_EXCEL_COLS)),
        max(1, min(r2, MAX_EXCEL_ROWS)),
        max(1, min(c2, MAX_EXCEL_COLS)),
    )


def shift_range_str(
    range_ref: str,
    row_ops: Sequence[AxisShiftOp],
    col_ops: Sequence[AxisShiftOp],
) -> str | None:
    """Shifts an A1 or A1:D10 range string, preserving '$' anchors. Returns None if deleted."""
    part1, sep, part2 = range_ref.strip().partition(":")
    r1, c1, c1_abs, r1_abs = parse_coord_anchored(part1)
    if not sep:
        new_r = shift_index(r1, row_ops)
        new_c = shift_index(c1, col_ops)
        if new_r is None or new_c is None:
            return None
        return format_coord(new_r, new_c, col_abs=c1_abs, row_abs=r1_abs)

    r2, c2, c2_abs, r2_abs = parse_coord_anchored(part2)
    shifted = shift_range_bounds(
        min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2), row_ops, col_ops
    )
    if shifted is None:
        return None
    nr1, nc1, nr2, nc2 = shifted
    return format_range(
        nr1,
        nc1,
        nr2,
        nc2,
        c1_abs=c1_abs,
        r1_abs=r1_abs,
        c2_abs=c2_abs,
        r2_abs=r2_abs,
    )


def shift_sqref(
    sqref: str,
    row_ops: Sequence[AxisShiftOp],
    col_ops: Sequence[AxisShiftOp],
) -> str:
    """Shifts space-separated sqref ranges (used in conditionalFormatting & dataValidations)."""
    out_parts: list[str] = []
    for token in sqref.split():
        shifted = shift_range_str(token, row_ops, col_ops)
        if shifted is not None:
            out_parts.append(shifted)
    return " ".join(out_parts)


def _normalize_sheet_name(raw_sheet: str) -> str:
    """Unquotes 'My Sheet' -> My Sheet and unescapes '' -> '."""
    if raw_sheet.startswith("'") and raw_sheet.endswith("'") and len(raw_sheet) >= 2:
        return raw_sheet[1:-1].replace("''", "'")
    return raw_sheet


def _try_parse_cell_token(token: str) -> tuple[int, int, bool, bool] | None:
    """Fast pure-character check for a cell token like B2 or $AA$100."""
    try:
        r, c, c_abs, r_abs = parse_coord_anchored(token)
        if r <= MAX_EXCEL_ROWS and c <= MAX_EXCEL_COLS:
            return r, c, c_abs, r_abs
    except ValueError:
        pass
    return None


def _try_parse_row_only(token: str) -> tuple[int, bool] | None:
    """Parses pure row token in row-range like '5' or '$5'."""
    if not token:
        return None
    abs_flag = token.startswith("$")
    digits = token[1:] if abs_flag else token
    if digits.isascii() and digits.isdigit():
        val = int(digits)
        if 1 <= val <= MAX_EXCEL_ROWS:
            return val, abs_flag
    return None


def _try_parse_col_only(token: str) -> tuple[int, bool] | None:
    """Parses pure column token in col-range like 'B' or '$AA'."""
    if not token:
        return None
    abs_flag = token.startswith("$")
    letters = token[1:] if abs_flag else token
    if 1 <= len(letters) <= 3 and letters.isascii() and letters.isalpha():
        val = col_to_idx(letters)
        if 1 <= val <= MAX_EXCEL_COLS:
            return val, abs_flag
    return None


def _rewrite_ref_token(
    ref_body: str,
    row_ops: Sequence[AxisShiftOp],
    col_ops: Sequence[AxisShiftOp],
) -> str:
    """Rewrites a cell, cell-range, row-range, or col-range token without regex."""
    if not row_ops and not col_ops:
        return ref_body

    lhs, sep, rhs = ref_body.partition(":")
    if not sep:
        cell = _try_parse_cell_token(lhs)
        if cell is None:
            return ref_body
        r, c, c_abs, r_abs = cell
        nr = shift_index(r, row_ops)
        nc = shift_index(c, col_ops)
        if nr is None or nc is None:
            return "#REF!"
        return format_coord(nr, nc, col_abs=c_abs, row_abs=r_abs)

    # Cell range: A1:B10
    c1 = _try_parse_cell_token(lhs)
    c2 = _try_parse_cell_token(rhs)
    if c1 is not None and c2 is not None:
        r1, col1, c1_abs, r1_abs = c1
        r2, col2, c2_abs, r2_abs = c2
        shifted = shift_range_bounds(
            min(r1, r2), min(col1, col2), max(r1, r2), max(col1, col2), row_ops, col_ops
        )
        if shifted is None:
            return "#REF!"
        nr1, nc1, nr2, nc2 = shifted
        return format_range(
            nr1,
            nc1,
            nr2,
            nc2,
            c1_abs=c1_abs,
            r1_abs=r1_abs,
            c2_abs=c2_abs,
            r2_abs=r2_abs,
        )

    # Whole-row range: 2:10 or $2:$10
    r_only1 = _try_parse_row_only(lhs)
    r_only2 = _try_parse_row_only(rhs)
    if r_only1 is not None and r_only2 is not None:
        rv1, r1_abs = r_only1
        rv2, r2_abs = r_only2
        shifted = shift_range_bounds(
            min(rv1, rv2), 1, max(rv1, rv2), 1, row_ops, ()
        )
        if shifted is None:
            return "#REF!"
        nr1, _, nr2, _ = shifted
        p1 = "$" if r1_abs else ""
        p2 = "$" if r2_abs else ""
        return f"{p1}{nr1}:{p2}{nr2}"

    # Whole-column range: B:D or $B:$D
    col_only1 = _try_parse_col_only(lhs)
    col_only2 = _try_parse_col_only(rhs)
    if col_only1 is not None and col_only2 is not None:
        cv1, c1_abs = col_only1
        cv2, c2_abs = col_only2
        shifted = shift_range_bounds(
            1, min(cv1, cv2), 1, max(cv1, cv2), (), col_ops
        )
        if shifted is None:
            return "#REF!"
        _, nc1, _, nc2 = shifted
        p1 = "$" if c1_abs else ""
        p2 = "$" if c2_abs else ""
        return f"{p1}{idx_to_col(nc1)}:{p2}{idx_to_col(nc2)}"

    return ref_body


def rewrite_formula(
    formula: str,
    current_sheet: str,
    shifts_by_sheet: Mapping[str, tuple[Sequence[AxisShiftOp], Sequence[AxisShiftOp]]],
    deleted_sheets: Iterable[str] = (),
) -> str:
    """
    Pure character-by-character scanner that rewrites cell/range references inside
    an Excel formula while ignoring string literals ("...") and function names.
    """
    deleted_set = set(deleted_sheets)
    out: list[str] = []
    n = len(formula)
    i = 0

    while i < n:
        ch = formula[i]

        # 1. Double-quoted string literal: copy verbatim (handling "" escapes)
        if ch == '"':
            start = i
            i += 1
            while i < n:
                if formula[i] == '"':
                    i += 1
                    if i < n and formula[i] == '"':
                        i += 1
                        continue
                    break
                i += 1
            out.append(formula[start:i])
            continue

        # 2. Bracketed structured ref (Table1[[#This Row],[Col]]) or external book ([1]Sheet1!A1)
        if ch == "[":
            start = i
            depth = 0
            while i < n:
                c = formula[i]
                if c == "'" and i + 1 < n:
                    i += 2  # ' escapes the next char inside structured refs
                    continue
                if c == "[":
                    depth += 1
                elif c == "]":
                    depth -= 1
                    if depth == 0:
                        i += 1
                        break
                i += 1
            # External workbook prefix: copy the following Sheet!Ref verbatim too
            j = i
            while j < n and (formula[j].isalnum() or formula[j] in ("_", ".")):
                j += 1
            if j > i and j < n and formula[j] == "!":
                j += 1
                while j < n and (formula[j].isalnum() or formula[j] in ("$", ":")):
                    j += 1
                i = j
            out.append(formula[start:i])
            continue

        # 3. Single-quoted sheet prefix: 'Sheet Name'!A1:B2
        if ch == "'":
            start = i
            i += 1
            while i < n:
                if formula[i] == "'":
                    i += 1
                    if i < n and formula[i] == "'":
                        i += 1
                        continue
                    break
                i += 1
            if i < n and formula[i] == "!":
                raw_prefix = formula[start:i]
                target_sheet = _normalize_sheet_name(raw_prefix)
                i += 1  # consume '!'
                ref_start = i
                while i < n and (
                    formula[i].isalnum() or formula[i] in ("$", ":")
                ):
                    i += 1
                ref_body = formula[ref_start:i]
                # Sheet names cannot contain '[', so this is an external workbook ref
                if "[" in target_sheet:
                    out.append(f"{raw_prefix}!{ref_body}")
                elif target_sheet in deleted_set:
                    out.append("#REF!")
                elif target_sheet in shifts_by_sheet:
                    r_ops, c_ops = shifts_by_sheet[target_sheet]
                    rewritten = _rewrite_ref_token(ref_body, r_ops, c_ops)
                    if rewritten == "#REF!":
                        out.append("#REF!")
                    else:
                        out.append(f"{raw_prefix}!{rewritten}")
                else:
                    out.append(f"{raw_prefix}!{ref_body}")
            else:
                out.append(formula[start:i])
            continue

        # 4. Unquoted identifier, sheet-qualified ref (Sheet2!A1), or local ref (A1:B5)
        if ch.isalpha() or ch in ("$", "_"):
            # Do not treat suffix after '.' (e.g. Table1.Column) as a cell coord
            if i > 0 and formula[i - 1] == ".":
                out.append(ch)
                i += 1
                continue

            start = i
            while i < n and (
                formula[i].isalnum() or formula[i] in ("$", "_", ".", ":")
            ):
                i += 1

            # Check if followed by '!' -> UnquotedSheet!Ref
            if i < n and formula[i] == "!":
                target_sheet = formula[start:i]
                i += 1  # consume '!'
                ref_start = i
                while i < n and (
                    formula[i].isalnum() or formula[i] in ("$", ":")
                ):
                    i += 1
                ref_body = formula[ref_start:i]
                if target_sheet in deleted_set:
                    out.append("#REF!")
                elif target_sheet in shifts_by_sheet:
                    r_ops, c_ops = shifts_by_sheet[target_sheet]
                    rewritten = _rewrite_ref_token(ref_body, r_ops, c_ops)
                    if rewritten == "#REF!":
                        out.append("#REF!")
                    else:
                        out.append(f"{target_sheet}!{rewritten}")
                else:
                    out.append(f"{target_sheet}!{ref_body}")
                continue

            token = formula[start:i]
            # Function call like SUM( or IF( -> leave untouched
            if i < n and formula[i] == "(":
                out.append(token)
                continue

            if current_sheet in shifts_by_sheet:
                r_ops, c_ops = shifts_by_sheet[current_sheet]
                out.append(_rewrite_ref_token(token, r_ops, c_ops))
            else:
                out.append(token)
            continue

        # 5. Pure numeric or row-range token (e.g. 2:10)
        if ch.isdigit():
            start = i
            while i < n and (formula[i].isdigit() or formula[i] in (".", ":", "$")):
                i += 1
            token = formula[start:i]
            if ":" in token and current_sheet in shifts_by_sheet:
                r_ops, c_ops = shifts_by_sheet[current_sheet]
                out.append(_rewrite_ref_token(token, r_ops, c_ops))
            else:
                out.append(token)
            continue

        out.append(ch)
        i += 1

    return "".join(out)
