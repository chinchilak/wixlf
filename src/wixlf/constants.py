from collections.abc import Callable
from datetime import date, datetime, time
from typing import TYPE_CHECKING, Any, Final

if TYPE_CHECKING:
    from .models import Formula

# OOXML Namespaces
NS_MAIN: Final[str] = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL_DOC: Final[str] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL_PKG: Final[str] = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT: Final[str] = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_XML: Final[str] = "http://www.w3.org/XML/1998/namespace"
NS_EXT_PROPS: Final[str] = (
    "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
)
NS_VT: Final[str] = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

APP_XML_PATH: Final[str] = "docProps/app.xml"

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
TAG_DV_FORMULA1: Final[str] = f"{{{NS_MAIN}}}formula1"
TAG_DV_FORMULA2: Final[str] = f"{{{NS_MAIN}}}formula2"
TAG_CF_RULE: Final[str] = f"{{{NS_MAIN}}}cfRule"
TAG_CF_FORMULA: Final[str] = f"{{{NS_MAIN}}}formula"
TAG_HYPERLINKS: Final[str] = f"{{{NS_MAIN}}}hyperlinks"
TAG_HYPERLINK: Final[str] = f"{{{NS_MAIN}}}hyperlink"
TAG_SHEET_VIEWS: Final[str] = f"{{{NS_MAIN}}}sheetViews"
TAG_SHEET_VIEW: Final[str] = f"{{{NS_MAIN}}}sheetView"
TAG_SHEET_FORMAT_PR: Final[str] = f"{{{NS_MAIN}}}sheetFormatPr"
TAG_PANE: Final[str] = f"{{{NS_MAIN}}}pane"
TAG_SELECTION: Final[str] = f"{{{NS_MAIN}}}selection"
TAG_ROW_BREAKS: Final[str] = f"{{{NS_MAIN}}}rowBreaks"
TAG_COL_BREAKS: Final[str] = f"{{{NS_MAIN}}}colBreaks"
TAG_BRK: Final[str] = f"{{{NS_MAIN}}}brk"
TAG_IGNORED_ERRORS: Final[str] = f"{{{NS_MAIN}}}ignoredErrors"
TAG_IGNORED_ERROR: Final[str] = f"{{{NS_MAIN}}}ignoredError"
TAG_PROTECTED_RANGES: Final[str] = f"{{{NS_MAIN}}}protectedRanges"
TAG_PROTECTED_RANGE: Final[str] = f"{{{NS_MAIN}}}protectedRange"
TAG_TABLE_PARTS: Final[str] = f"{{{NS_MAIN}}}tableParts"
TAG_TABLE_PART: Final[str] = f"{{{NS_MAIN}}}tablePart"
TAG_TABLE: Final[str] = f"{{{NS_MAIN}}}table"
TAG_TABLE_COLUMNS: Final[str] = f"{{{NS_MAIN}}}tableColumns"
TAG_TABLE_COLUMN: Final[str] = f"{{{NS_MAIN}}}tableColumn"
TAG_CALC_COL_FORMULA: Final[str] = f"{{{NS_MAIN}}}calculatedColumnFormula"
TAG_TOTALS_ROW_FORMULA: Final[str] = f"{{{NS_MAIN}}}totalsRowFormula"
TAG_FILTER_COLUMN: Final[str] = f"{{{NS_MAIN}}}filterColumn"
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
type CellScalar = str | int | float | bool | date | datetime | time | Formula | None
type PatchWorkerFn = Callable[[str, str, Any, str], None]
