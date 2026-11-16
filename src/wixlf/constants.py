from collections.abc import Callable
from datetime import date, datetime, time
from typing import Any, Final

# OOXML Namespaces
NS_MAIN: Final[str] = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL_DOC: Final[str] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL_PKG: Final[str] = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT: Final[str] = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_XML: Final[str] = "http://www.w3.org/XML/1998/namespace"

# Pre-qualified SpreadsheetML element tags
TAG_WORKSHEET: Final[str] = f"{{{NS_MAIN}}}worksheet"
TAG_SHEET_DATA: Final[str] = f"{{{NS_MAIN}}}sheetData"
TAG_DIMENSION: Final[str] = f"{{{NS_MAIN}}}dimension"
TAG_ROW: Final[str] = f"{{{NS_MAIN}}}row"
TAG_CELL: Final[str] = f"{{{NS_MAIN}}}c"
TAG_VALUE: Final[str] = f"{{{NS_MAIN}}}v"
TAG_FORMULA: Final[str] = f"{{{NS_MAIN}}}f"
TAG_INLINE_STR: Final[str] = f"{{{NS_MAIN}}}is"
TAG_TEXT: Final[str] = f"{{{NS_MAIN}}}t"

# Calculation chain identifiers
CALC_CHAIN_PATH: Final[str] = "xl/calcChain.xml"
CALC_CHAIN_CT: Final[str] = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"
)
CALC_CHAIN_REL_TYPE: Final[str] = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain"
)

# PEP 695 native type aliases
type CellScalar = str | int | float | bool | date | datetime | time | None
type PatchWorkerFn = Callable[[str, str, Any, str], None]
