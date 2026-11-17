# WIXLF — Write Into XLSX Files

A high-performance, constant-memory Python 3.14+ library for modifying existing Excel (`.xlsx` / `.xlsm`) workbooks **in-place**. Instead of deserializing the entire workbook DOM into memory like `openpyxl`, **WIXLF** streams modified and newly created worksheets row-by-row via `lxml` while copying untouched archive parts byte-for-byte.

---

## Key Features

* **In-Place Modification by Default:** Updates your existing workbook directly via atomic staging (`os.replace()`), with an optional `dst_path` parameter whenever you want to save to a separate copy.
* **Worksheet Tab Management:** Create new populated or blank sheets (at any index or appended to the end), delete worksheets, and reorder sheet tabs (`move_sheet` to an index or to the end).
* **Constant $O(1)$ Memory Footprint:** Streams `<worksheet>` XML via `lxml.etree.iterparse` + `lxml.etree.xmlfile` with immediate DOM node pruning (`elem.clear()` and preceding-sibling deletion), keeping RAM flat even on multi-hundred-megabyte sheets.
* **Per-Worksheet Multiprocessing:** Dispatches dirty and newly created worksheets across a `ProcessPoolExecutor` (bypassing the GIL) when handling multiple sheets or sheets $\ge$ 512 KB, while running single small edits inline to avoid process spawn overhead.
* **Lossless Workbook Round-Tripping:** Untouched worksheets, `styles.xml`, themes, charts, pivot tables, drawings, media, and VBA macros (`.xlsm`) are streamed straight from source to destination in 1 MB chunks without parsing or stripping unsupported OOXML tags.
* **Automatic Style Inheritance:** Overwriting an existing cell preserves its original Excel style index (`s` attribute—fonts, fills, borders, number formats) unless explicitly overridden or cleared.
* **Zero `sharedStrings.xml` Contention:** Writes text values using OOXML `inlineStr` (`<c t="inlineStr"><is><t>...</t></is></c>`), eliminating shared-string table locks and enabling safe parallel sheet processing.
* **Automatic Formula Recalculation Safety:** Strips stale `xl/calcChain.xml` references and sets `fullCalcOnLoad="1"` in `xl/workbook.xml` so Excel recalculates dependent formulas cleanly on open.
* **Pure-Python Utilities (Zero Regex):** Cell coordinate parsing and XML 1.0 character sanitization use pure ASCII character scanning and C-level `str.translate()`.

---

## Installation

Requires **Python 3.14+** and `lxml`:

```bash
pip install wixlf
```

---

## Quick Start

Only `import wixlf` is required—all functions and classes are available at the package root.

### 1. Modify Existing Sheets In-Place

```python
import wixlf

# Modifies report.xlsx directly (pass dst_path="copy.xlsx" to save elsewhere)
wixlf.patch_workbook(
    "report.xlsx",
    {
        "Sales": {
            "A1": "Q4 Revenue Report",
            "B2": 149500.75,
            "C2": True,
            "D2": "=B2*0.21",
        },
        "Summary": {
            (2, 1): "Total",
            (2, 2): "=SUM(Sales!B2:B500)",
        },
    },
)
```

### 2. Create, Delete, and Reorder Worksheets

```python
import wixlf

# Add a new sheet at the end (default when position=None)
wixlf.add_sheet(
    "report.xlsx",
    name="Audit Log",
    data={"A1": "Timestamp", "B1": "Status", "A2": "2026-11-17", "B2": "OK"},
)

# Add a new sheet at the very front (0-based index)
wixlf.add_sheet("report.xlsx", name="Cover", data={"A1": "Executive Summary"}, position=0)

# Move an existing sheet to index 1 (second tab)
wixlf.move_sheet("report.xlsx", "Summary", position=1)

# Move a sheet to the end (position=None)
wixlf.move_sheet("report.xlsx", "Cover", position=None)

# Delete one or multiple sheets
wixlf.delete_sheets("report.xlsx", ["OldData", "Temp"])
```

### 3. Combined Atomic Operations in One Pass

You can combine cell edits, new sheets, deletions, and tab moves into a single `patch_workbook()` call so the archive is only rewritten once:

```python
import wixlf

wixlf.patch_workbook(
    "report.xlsx",
    edits_by_sheet={
        "Sales": {"B2": 162000.00},
    },
    new_sheets=[
        wixlf.NewSheetSpec(
            name="Q4 Export",
            plan=wixlf.SheetEditPlan.from_dict({"A1": "Exported", "A2": 162000.00}),
            position=None,  # Appends to end
        ),
    ],
    delete_sheets=["Scratch"],
    move_sheets=[
        wixlf.SheetMoveSpec(name="Summary", position=0),   # Move to first tab
        wixlf.SheetMoveSpec(name="Q4 Export", position=None),  # Move to end
    ],
    dst_path=None,  # None = modify report.xlsx in-place
)
```

---

## Supported Cell Types & Styling

### Data Types & Formulas

`SheetEditPlan.from_dict()` accepts A1-style coordinates (`"B5"`, `"$AA$10"`) or 1-based `(row, col)` tuples (`(5, 2)`):

```python
from datetime import date, datetime, time
import wixlf

plan = wixlf.SheetEditPlan.from_dict({
    "A2": "Plain text (auto-strips illegal XML 1.0 control chars)",
    "B2": "  Leading/trailing spaces (auto-sets xml:space='preserve')  ",
    "C2": 42,                        # Integer
    "D2": 3.14159,                   # Float (NaN/Inf serializes as #NUM!)
    "E2": True,                      # Boolean (OOXML t="b")
    "F2": date(2026, 11, 17),        # Converted to Excel 1900 serial number
    "G2": datetime(2026, 11, 17, 9, 30),
    "H2": time(9, 30, 0),
    "I2": "=SUM(C2:D2)",              # Formula (leading '=' stripped for OOXML <f>)
    "J2": None,                      # Clears cell value while keeping cell style
})
```

### Overriding or Clearing Cell Styles

```python
import wixlf

# Option A: Assign explicit style indices ('s' from xl/styles.xml)
plan = wixlf.SheetEditPlan.from_dict(
    cell_map={"A10": 9999.95, "B10": "Total"},
    styles={"A10": "14", "B10": "15"},
)

# Option B: Strip styling from specific cells
custom_plan = wixlf.SheetEditPlan(
    rows=(
        wixlf.RowEditBatch(
            row=10,
            cells=(
                wixlf.CellEdit(row=10, col=1, value="Unstyled Reset", clear_style=True),
            ),
        ),
    )
)
```

---

## Package Architecture

```text
src/wixlf/
├── __init__.py          # Public API exports: patch_workbook, add_sheet, delete_sheets, move_sheet
├── constants.py         # OOXML namespaces, QName tags, content-types, and type aliases
├── models.py            # Immutable slotted dataclasses (SheetInfo, CellEdit, SheetEditPlan, NewSheetSpec, SheetMoveSpec)
├── reader.py            # PackageReader: resolves sheet names & tab order to internal zip XML paths
├── patcher.py           # SheetStreamPatcher, patch_sheet_worker & create_sheet_worker (O(1) memory)
├── assembler.py         # PackageAssembler: ProcessPoolExecutor orchestrator, sheet operations & atomic zip writer
└── utils/
    ├── __init__.py      # Re-exports shared utility functions
    ├── coords.py        # Pure-Python coordinate/column converters (col_to_idx, parse_coord, etc.)
    ├── xml_text.py      # XML 1.0 text sanitization & Excel 1900 serial date conversion
    └── zip_io.py        # Streaming zip entry helpers, ZipInfo cloning, and OOXML path resolution
```

---

## Current Limitations

* **No Row/Column Shifting:** Does not insert/delete rows inside an existing sheet (shifting downstream formulas) or auto-resize Excel Tables (`xl/tables/tableN.xml`).
* **Date Cells Require Existing Date Formatting:** `date` and `datetime` values serialize to Excel serial numbers; the target cell must already have a date number format style (`s`) in the template to render as a date rather than a number.
* **No New Style Creation:** Can inherit or assign existing `s` style indices from the workbook, but does not author new entries in `xl/styles.xml`.
