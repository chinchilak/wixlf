# WIXLF — Write Into XLSX Files

A high-performance, constant-memory Python 3.14+ library for modifying existing Excel (`.xlsx` / `.xlsm`) workbooks. Instead of deserializing the entire workbook DOM into memory like `openpyxl`, **WIXLF** streams modified worksheets row-by-row via `lxml` while copying untouched archive parts byte-for-byte.

#### Why openpyxl Is Slow & Limiting for Existing Files
If you are designing a streamlined writer specifically for modifying existing files, openpyxl suffers from four architectural bottlenecks you can avoid:

* Full DOM Deserialization: Even to change a single cell (ws['A1'] = 5), openpyxl parses every XML file in the .xlsx zip archive (sharedStrings.xml, styles.xml, all sheetN.xml files) into heavy Python objects and re-serializes everything on save(). A streaming XML/zip-patching approach that only rewrites the modified sheetN.xml (and calcChain.xml if needed) is orders of magnitude faster.

* No write_only Mode for Existing Files: openpyxl restricts its fast lxml-backed streaming writer (write_only=True) to brand-new workbooks. You cannot stream-append rows to an existing template without loading the whole file into memory.

* Destructive Round-Tripping: Because openpyxl only rebuilds XML elements it explicitly models, opening and saving an existing workbook silently strips unsupported features—such as shapes, SmartArt, form controls, slicers, and certain pivot table/chart properties. Preserving untouched XML parts verbatim inside the zip archive avoids this completely.

* No Automatic Formula Shifting: Inserting/deleting rows or appending data into existing Excel Tables (ListObject in xl/tables/tableN.xml) requires manual updates to table ranges and formulas in openpyxl.

---

## Key Features

* **Constant $O(1)$ Memory Footprint:** Streams `<worksheet>` XML via `lxml.etree.iterparse` + `lxml.etree.xmlfile` with immediate DOM node pruning (`elem.clear()` and preceding-sibling deletion), keeping RAM flat even on multi-hundred-megabyte sheets.
* **Per-Worksheet Multiprocessing:** Automatically dispatches modified worksheets across a `ProcessPoolExecutor` (bypassing the GIL) when editing multiple sheets or sheets $\ge$ 512 KB, while running single small edits inline to avoid process spawn overhead.
* **Lossless Workbook Round-Tripping:** Untouched worksheets, `styles.xml`, themes, charts, pivot tables, drawings, media, and VBA macros (`.xlsm`) are streamed straight from source to destination in 1 MB chunks without parsing or stripping unsupported OOXML tags.
* **Automatic Style Inheritance:** Overwriting an existing cell preserves its original Excel style index (`s` attribute—fonts, fills, borders, number formats) unless explicitly overridden or cleared.
* **Zero `sharedStrings.xml` Contention:** Writes text values using OOXML `inlineStr` (`<c t="inlineStr"><is><t>...</t></is></c>`), eliminating shared-string table locks and enabling safe parallel sheet processing.
* **Automatic Formula Recalculation Safety:** Strips stale `xl/calcChain.xml` references from `[Content_Types].xml` and `xl/_rels/workbook.xml.rels` and sets `fullCalcOnLoad="1"` in `xl/workbook.xml` so Excel recalculates dependent formulas cleanly on open.
* **Atomic File Replacement:** Writes to a staged temporary `.xlsx` file in the target directory and promotes it via `os.replace()`, preventing half-written or corrupted outputs on failure.
* **Pure-Python Utilities (Zero Regex):** Cell coordinate parsing and XML 1.0 character sanitization use pure ASCII character scanning and C-level `str.translate()`.

---

## Installation

Requires **Python 3.14+** and `lxml`:

```bash
pip install wixlf
```

Or install from source:

```bash
pip install -e .
```

---

## Quick Start

### High-Level API (`wixlf.patch_workbook`)

```python
import wixlf

wixlf.patch_workbook(
    "template.xlsx",
    "output.xlsx",
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

### Lower-Level Pipeline API

Use `PackageReader`, `SheetEditPlan`, and `PackageAssembler` directly when you need fine-grained control over worker count, chunk sizes, or pre-compiled edit plans:

```python
from wixlf import (
    PackageAssembler,
    PackageReader,
    SheetEditPlan,
    patch_sheet_worker,
)

# 1. Resolve human-readable sheet names to internal zip paths
reader = PackageReader("template.xlsx")
sheets = reader.get_sheets()

# 2. Pre-sort edits into immutable, picklable plans per sheet
dirty_edits = {
    sheets["Sales"].zip_path: SheetEditPlan.from_dict({
        "A1": "Q4 Revenue Report",
        "B2": 149500.75,
    }),
}

# 3. Execute parallel sheet streaming & atomic archive assembly
assembler = PackageAssembler(
    src_path="template.xlsx",
    dst_path="output.xlsx",
    max_workers=4,               # Defaults to os.process_cpu_count()
    chunk_size=2 * 1024 * 1024,  # 2 MB zip stream copy buffer
)
assembler.build(
    dirty_edits=dirty_edits,
    patch_fn=patch_sheet_worker,
    force_recalc=True,
)
```

---

## Supported Cell Types & Styling

### 1. Data Types & Formulas

`SheetEditPlan.from_dict()` accepts A1-style coordinates (`"B5"`, `"$AA$10"`) or 1-based `(row, col)` tuples (`(5, 2)`):

```python
from datetime import date, datetime, time
from wixlf import SheetEditPlan

plan = SheetEditPlan.from_dict({
    "A2": "Plain text (auto-strips illegal XML 1.0 control chars)",
    "B2": "  Leading/trailing spaces (auto-sets xml:space='preserve')  ",
    "C2": 42,                        # Integer
    "D2": 3.14159,                   # Float (NaN/Inf serializes as #NUM!)
    "E2": True,                      # Boolean (OOXML t="b")
    "F2": date(2026, 11, 16),        # Converted to Excel 1900 serial number
    "G2": datetime(2026, 11, 16, 14, 30),
    "H2": time(9, 15, 0),
    "I2": "=SUM(C2:D2)",              # Formula (leading '=' stripped for OOXML <f>)
    "J2": None,                      # Clears cell value while keeping cell style
})
```

### 2. Overriding or Clearing Cell Styles

```python
from wixlf import CellEdit, RowEditBatch, SheetEditPlan

# Option A: Assign explicit style indices ('s' from xl/styles.xml)
plan = SheetEditPlan.from_dict(
    cell_map={"A10": 9999.95, "B10": "Total"},
    styles={"A10": "14", "B10": "15"},
)

# Option B: Strip styling from specific cells
custom_plan = SheetEditPlan(
    rows=(
        RowEditBatch(
            row=10,
            cells=(
                CellEdit(row=10, col=1, value="Unstyled Reset", clear_style=True),
            ),
        ),
    )
)
```

---

## Package Architecture

```text
src/wixlf/
├── __init__.py          # Public API exports + patch_workbook() facade
├── constants.py         # OOXML namespaces, QName tags, content-types, and type aliases
├── models.py            # Immutable slotted dataclasses (SheetInfo, CellEdit, RowEditBatch, SheetEditPlan)
├── reader.py            # PackageReader: resolves sheet names to internal zip XML paths
├── patcher.py           # SheetStreamPatcher & patch_sheet_worker: O(1) memory XML stream patcher
├── assembler.py         # PackageAssembler: ProcessPoolExecutor orchestrator & atomic zip writer
└── utils/
    ├── __init__.py      # Re-exports shared utility functions
    ├── coords.py        # Pure-Python coordinate/column converters (col_to_idx, parse_coord, etc.)
    ├── xml_text.py      # XML 1.0 text sanitization & Excel 1900 serial date conversion
    └── zip_io.py        # Streaming zip entry helpers, ZipInfo cloning, and OOXML path resolution
```

---

## Current MVP Limitations

* **Cell Value & Formula Updates Only:** Does not insert/delete rows (shifting formulas) or auto-resize Excel Tables (`xl/tables/tableN.xml`).
* **Date Cells Require Existing Date Formatting:** `date` and `datetime` values serialize to Excel serial numbers; the target cell must already have a date number format style (`s`) in the template to render as a date rather than a number.
* **No New Style Creation:** Can inherit or assign existing `s` style indices from the workbook, but does not author new entries in `xl/styles.xml`.
