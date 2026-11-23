# WIXLF — Write Into XLSX Files

A high-performance, constant-memory Python 3.14+ library for modifying existing Excel (`.xlsx` / `.xlsm`) workbooks **in-place**. Instead of deserializing the entire workbook DOM into memory like `openpyxl`, **WIXLF** streams modified and newly created worksheets row-by-row via `lxml` and transfers archive entries in fixed-size chunks.

---

## Key Features

* **In-Place Modification by Default:** Updates your existing workbook directly via atomic staging (`os.replace()`), with an optional `dst_path` parameter whenever you want to save to a separate copy.
* **Row & Column Insertion / Deletion with Formula & Table Shifting:** Insert or delete rows and columns (`insert_rows`, `delete_rows`, `insert_cols`, `delete_cols`) while automatically shifting downstream cells, cross-sheet formulas (`<f>`), Excel Tables (`xl/tables/tableN.xml`), `<mergeCells>`, `<cols>` widths, `<autoFilter>`, conditional formatting, data validations, and `<definedNames>`.
* **Clean Cell & Range Clearing (`clear_cells`):** Wipe cell values, strip styles (`clear_style=True`), or remove `<c>` elements (`drop_cells=True`) across individual coordinates or `A1:D10` ranges without shifting rows/columns.
* **Worksheet Tab Management:** Create new populated or blank sheets (at any index or appended to the end), delete worksheets, and reorder sheet tabs (`move_sheet` to an index or to the end).
* **Bounded-Memory ZIP Transfer:** Archive entries are copied one at a time in 1 MB chunks. Unchanged DEFLATE entries are recompressed using the selected compression level; package parts such as styles, charts, media, and VBA macros (`.xlsm`) are preserved.
* **Fast Deflate (`compresslevel=1`):** Modified and newly created XML entries default to `compresslevel=1` (configurable up to `9`), accelerating writes while keeping file sizes compact.
* **Constant $O(1)$ Memory Footprint:** Streams `<worksheet>` XML via `lxml.etree.iterparse` + `lxml.etree.xmlfile` with immediate DOM node pruning (`elem.clear()` and preceding-sibling deletion).
* **Per-Worksheet Multiprocessing:** Dispatches dirty and newly created worksheets across a `ProcessPoolExecutor` (bypassing the GIL) when 2+ sheets are processed and their combined size is $\ge$ 512 KB (tunable via `PackageAssembler(parallel_threshold_bytes=...)`); smaller jobs run inline.
* **Pure-Python Utilities (Zero Regex):** Coordinate parsing, formula reference rewriting, and XML 1.0 character sanitization use pure character-by-character scanning and C-level `str.translate()`.

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
            "D2": wixlf.Formula("=B2*0.21"),
        },
        "Summary": {
            (2, 1): "Total",
            (2, 2): wixlf.Formula("=SUM(Sales!B2:B500)"),
        },
    },
)
```

### 2. Insert, Delete & Clear Rows, Columns, and Cells

```python
import wixlf

# Insert 3 blank rows at row 5 (shifts downstream cells, formulas & Excel Tables down)
wixlf.insert_rows("report.xlsx", "Sales", row_idx=5, count=3)

# Delete 2 rows starting at row 10 (formulas pointing to deleted cells become #REF!)
wixlf.delete_rows("report.xlsx", "Sales", row_idx=10, count=2)

# Insert 1 column at column 'C' (or 1-based index 3)
wixlf.insert_cols("report.xlsx", "Sales", col="C", count=1)

# Delete 2 columns starting at column 'E'
wixlf.delete_cols("report.xlsx", "Sales", col="E", count=2)

# Clear values in a range while keeping cell formatting
wixlf.clear_cells("report.xlsx", "Sales", "B2:D20")

# Wipe values + styling, or remove <c> XML elements completely
wixlf.clear_cells("report.xlsx", "Sales", ["A1", "B2:C5"], clear_style=True, drop_cells=True)
```

### 3. Create, Delete, and Reorder Worksheets

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

### 4. Show and Hide Worksheets, Rows, and Columns

```python
import wixlf

# Unhide every row and column in every worksheet
wixlf.unhide_all_rows_and_columns("report.xlsx")

# Unhide every worksheet tab, including veryHidden sheets
wixlf.unhide_all_sheets("report.xlsx")

# Hide rows and columns by worksheet (indices are 1-based; columns accept letters or indices)
wixlf.set_visibility(
    "report.xlsx",
    rows={"Sales": [2, 5], "Summary": [10]},
    columns={"Sales": ["B", "D", "E", "F"]},
    hidden=True,
)

# Unhide selected rows and columns
wixlf.set_visibility(
    "report.xlsx",
    rows={"Sales": [5]},
    columns={"Sales": ["D"]},
    hidden=False,
)

# Worksheet tabs support Excel's visible, hidden, and veryHidden states
wixlf.set_visibility("report.xlsx", sheets={"Summary": "veryHidden"})
wixlf.set_visibility("report.xlsx", sheets={"Summary": "visible"})
```

For rows and columns, use `hidden=True` to hide the listed targets and `hidden=False` to unhide them. Use `"all"` as a per-sheet target only with `hidden=False`. `unhide_all_rows_and_columns()` also processes hidden and `veryHidden` sheets without changing their tab visibility. Worksheet tab states are set independently via the `sheets` mapping. `unhide_all_sheets()` leaves row and column visibility unchanged. The workbook must retain at least one visible worksheet.

### 5. Combined Atomic Operations in One Pass

Combine cell edits, row/column shifts, new sheets, deletions, and tab moves into a single `patch_workbook()` call so the archive is only rewritten once:

```python
import wixlf

wixlf.patch_workbook(
    "report.xlsx",
    edits_by_sheet={
        "Sales": {"B5": 162000.00},
    },
    shift_ops={
        "Sales": (
            [wixlf.AxisShiftOp(start=5, delta=1)],   # Insert 1 row at row 5
            [wixlf.AxisShiftOp(start=3, delta=-1)],  # Delete column C
        ),
    },
    new_sheets=[
        wixlf.NewSheetSpec(
            name="Q4 Export",
            plan=wixlf.SheetEditPlan.from_dict({"A1": "Exported", "A2": 162000.00}),
            position=None,
        ),
    ],
    delete_sheets=["Scratch"],
    move_sheets=[
        wixlf.SheetMoveSpec(name="Summary", position=0),
        wixlf.SheetMoveSpec(name="Q4 Export", position=None),
    ],
    compresslevel=1,
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
    "I2": wixlf.Formula("=SUM(C2:D2)"),  # Formula (leading '=' optional)
    "K2": "=not a formula",          # Plain str is always text (formula-injection safe)
    "J2": None,                      # Clears cell value while keeping cell style
})
```

To restore the legacy behaviour where any `str` starting with `=` is a formula, pass `formulas_from_strings=True` to `patch_workbook()` / `add_sheet()`.

### Overriding or Clearing Cell Styles

```python
import wixlf

# Option A: Assign explicit style indices ('s' from xl/styles.xml)
plan = wixlf.SheetEditPlan.from_dict(
    cell_map={"A10": 9999.95, "B10": "Total"},
    styles={"A10": "14", "B10": "15"},
)

# Option B: Strip styling or drop <c> elements from specific cells
custom_plan = wixlf.SheetEditPlan(
    rows=(
        wixlf.RowEditBatch(
            row=10,
            cells=(
                wixlf.CellEdit(row=10, col=1, value="Unstyled Reset", clear_style=True),
                wixlf.CellEdit(row=10, col=2, value=None, drop_cell=True),
            ),
        ),
    )
)
```

### Windows / macOS: guard your entry point

Worker processes are started with `spawn`, which re-imports your script. Always call wixlf from behind a main guard:

```python
import wixlf

if __name__ == "__main__":
    wixlf.patch_workbook("report.xlsx", {"Sales": {"A1": "Hello"}})
```

---

## Package Architecture

```text
.
├── .gitignore
├── .github/
│   └── workflows/
│       └── publish.yml  # Package publishing workflow
├── README.md
├── pyproject.toml       # Project metadata, dependencies, and build configuration
├── requirements.txt
├── test_readme.py       # Root-level unittest coverage for documented behavior
└── src/
    └── wixlf/
        ├── __init__.py  # Package-root API and convenience functions
        ├── assembler.py # Workbook ZIP assembly, table/defined-name updates, atomic output
        ├── constants.py # OOXML namespaces, element tags, and shared type aliases
        ├── models.py    # Immutable edit plans, worksheet metadata, and operation payloads
        ├── patcher.py   # Streaming worksheet XML patcher and worker entry points
        ├── reader.py    # Reads workbook metadata and maps sheet names to archive paths
        ├── visibility.py # Worksheet-tab, row, and column visibility operations
        └── utils/
            ├── __init__.py # Shared utility exports
            ├── coords.py   # Cell/range/column coordinate parsing and formatting
            ├── names.py    # Worksheet-name validation
            ├── shifting.py # Row/column index, range, and formula-reference shifting
            ├── xml_text.py # Hardened XML parsing, text sanitization, and date serials
            └── zip_io.py   # ZIP entry streaming, ZipInfo cloning, and relationship helpers
```

---

## Current Limitations

* **Date Cells Require Existing Date Formatting:** `date` and `datetime` values serialize to Excel serial numbers; the target cell must already have a date number format style (`s`) in the template to render as a date rather than a number.
* **No New Style Creation:** Can inherit or assign existing `s` style indices from the workbook, but does not author new entries in `xl/styles.xml`.
* **Cross-Sheet Dependents of Deleted Sheets:** Deleting a worksheet removes its own drawings, comments, tables, and other parts, but pivot tables, pivot caches, or charts on *other* sheets that source data from it are not rewritten; Excel will repair or empty them on open.
