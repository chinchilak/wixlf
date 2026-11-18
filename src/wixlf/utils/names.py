from typing import Final

MAX_SHEET_NAME_LEN: Final[int] = 31
_FORBIDDEN_SHEET_CHARS: Final[frozenset[str]] = frozenset(":\\/?*[]")


def validate_sheet_name(name: str) -> None:
    """Raises ValueError if `name` violates Excel's worksheet naming rules."""
    if not name:
        raise ValueError("Worksheet name cannot be empty.")
    if len(name) > MAX_SHEET_NAME_LEN:
        raise ValueError(
            f"Worksheet name {name!r} exceeds {MAX_SHEET_NAME_LEN} characters."
        )
    bad_chars = sorted({ch for ch in name if ch in _FORBIDDEN_SHEET_CHARS})
    if bad_chars:
        raise ValueError(
            f"Worksheet name {name!r} contains forbidden character(s): {''.join(bad_chars)}"
        )
    if name.startswith("'") or name.endswith("'"):
        raise ValueError(f"Worksheet name {name!r} cannot start or end with an apostrophe.")
    if name.casefold() == "history":
        raise ValueError("'History' is a reserved worksheet name.")
