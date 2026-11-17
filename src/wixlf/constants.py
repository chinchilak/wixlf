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
TAG_COLS: Final[str] = f"{{{NS_MAIN}}}cols"
TAG_COL: Final[str] = f"{{{NS_MAIN}}}col"
TAG_ROW: Final[str] = f"{{{NS_MAIN}}}row"
TAG_CELL: Final[str] = f"{{{NS_MAIN}}}c"
TAG_VALUE: Final[str] = f"{{{NS_MAIN}}}v"
TAG_FORMULA: Final[str] = f"{{{NS_MAIN}}}f"
TAG_INLINE_STR: Final[str] = f"{{{NS_MAIN}}}is"
TAG_TEXT: Final[str] = f"{{{NS_MAIN}}}t"
TAG_MERGE_CELLS: Final[str] = f"{{{NS_MAIN}}}mergeCells"
TAG_MERGE_CELL: Final[str] = f"{{{NS_MAIN}}}mergeCell"
TAG_AUTO_FILTER: Final[str] = f"{{{NS_MAIN}}}autoFilter"
TAG_COND_FMT: Final[str] = f"{{{NS_MAIN}}}conditionalFormatting"
TAG_DATA_VALS: Final[str] = f"{{{NS_MAIN}}}dataValidations"
TAG_DATA_VAL: Final[str] = f"{{{NS_MAIN}}}dataValidation"
TAG_TABLE_PARTS: Final[str] = f"{{{NS_MAIN}}}tableParts"
TAG_TABLE_PART: Final[str] = f"{{{NS_MAIN}}}tablePart"
TAG_TABLE: Final[str] = f"{{{NS_MAIN}}}table"
TAG_DEFINED_NAMES: Final[str] = f"{{{NS_MAIN}}}definedNames"
TAG_DEFINED_NAME: Final[str] = f"{{{NS_MAIN}}}definedName"

# Content Types & Relationship Types
WORKSHEET_CT: Final[str] = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
)
WORKSHEET_REL_TYPE: Final[str] = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
)
TABLE_REL_TYPE: Final[str] = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"
)

CALC_CHAIN_PATH: Final[str] = "xl/calcChain.xml"
CALC_CHAIN_CT: Final[str] = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"
)
CALC_CHAIN_REL_TYPE: Final[str] = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain"
)

# Excel grid limits (XFD = 16384, max row = 1048576)
MAX_EXCEL_ROWS: Final[int] = 1_048_576
MAX_EXCEL_COLS: Final[int] = 16_384

# PEP 695 native type aliases
type CellScalar = str | int | float | bool | date | datetime | time | None
type PatchWorkerFn = Callable[[str, str, Any, str], None]
