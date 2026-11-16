# WIXLF - Write Into XLsx Files

A high-performance, constant-memory Python 3.14+ engine for modifying existing Excel (`.xlsx` / `.xlsm`) workbooks. Instead of deserializing the entire workbook DOM into memory like `openpyxl`, it streams modified worksheets row-by-row via `lxml` while copying untouched archive parts byte-for-byte.

#### Why openpyxl Is Slow & Limiting for Existing Files
If you are designing a streamlined writer specifically for modifying existing files, openpyxl suffers from four architectural bottlenecks you can avoid:

* Full DOM Deserialization: Even to change a single cell (ws['A1'] = 5), openpyxl parses every XML file in the .xlsx zip archive (sharedStrings.xml, styles.xml, all sheetN.xml files) into heavy Python objects and re-serializes everything on save(). A streaming XML/zip-patching approach that only rewrites the modified sheetN.xml (and calcChain.xml if needed) is orders of magnitude faster.

* No write_only Mode for Existing Files: openpyxl restricts its fast lxml-backed streaming writer (write_only=True) to brand-new workbooks. You cannot stream-append rows to an existing template without loading the whole file into memory.

* Destructive Round-Tripping: Because openpyxl only rebuilds XML elements it explicitly models, opening and saving an existing workbook silently strips unsupported features—such as shapes, SmartArt, form controls, slicers, and certain pivot table/chart properties. Preserving untouched XML parts verbatim inside the zip archive avoids this completely.

* No Automatic Formula Shifting: Inserting/deleting rows or appending data into existing Excel Tables (ListObject in xl/tables/tableN.xml) requires manual updates to table ranges and formulas in openpyxl.

---

## Features

* **Constant $O(1)$ Memory Footprint:** Uses `lxml.etree.iterparse` paired with `lxml.etree.xmlfile` and aggressive sibling pruning (`elem.clear()` + removing preceding nodes) so memory stays flat even on multi-hundred-megabyte sheets.
* **Per-Worksheet Multiprocessing:** Automatically dispatches modified worksheets across a `ProcessPoolExecutor` (bypassing the GIL) when editing multiple sheets or sheets $\ge$ 512 KB, while running single small edits inline to avoid process spawn overhead.
* **Lossless Workbook Round-Tripping:** Untouched worksheets, `styles.xml`, themes, charts, pivot tables, drawings, media, and VBA macros (`.xlsm`) are streamed straight from source to destination in 1 MB chunks without parsing or stripping unsupported OOXML tags.
* **Automatic Style Inheritance:** Overwriting an existing cell automatically preserves its original Excel style index (`s` attribute—fonts, fills, borders, number formats) unless explicitly overridden or cleared.
* **Zero `sharedStrings.xml` Contention:** Writes text values using OOXML `inlineStr` (`<c t="inlineStr"><is><t>...</t></is></c>`), eliminating shared-string table locks and enabling safe parallel sheet processing.
* **Automatic Formula Recalculation Safety:** Strips stale `xl/calcChain.xml` references from `[Content_Types].xml` and `xl/_rels/workbook.xml.rels` and sets `fullCalcOnLoad="1"` in `xl/workbook.xml` so Excel recalculates dependent formulas cleanly on open.
* **Atomic File Replacement:** Writes to a staged temporary `.xlsx` file in the target directory and promotes it via `os.replace()`, preventing half-written or corrupted outputs on failure.

---

## Requirements

* **Python:** 3.14+ (uses PEP 695 `type` statements, native `collections.abc` generics, `os.process_cpu_count()`, and slotted dataclasses)
* **Dependencies:** `lxml` (see `requirements.txt`)

```bash
pip install -r requirements.txt
```

---

## Quick Start

```python
from package_pipeline import PackageAssembler, PackageReader
from sheet_patcher import SheetEditPlan, patch_sheet_worker

# 1. Inspect sheet names and resolve internal archive paths
reader = PackageReader("template.xlsx")
sheets = reader.get_sheets()

# 2. Compile sorted edit plans per target worksheet
dirty_edits = {
    sheets["Sales"].zip_path: SheetEditPlan.from_dict({
        "A1": "Q4 Revenue Report",
        "B2": 149500.75,
        "C2": True,
        "D2": "=B2*0.21",
    }),
    sheets["Summary"].zip_path: SheetEditPlan.from_dict({
        (2, 1): "Total",
        (2, 2): "=SUM(Sales!B2:B500)",
    }),
}

# 3. Stream-patch dirty sheets in parallel and assemble the output workbook
assembler = PackageAssembler("template.xlsx", "output.xlsx")
assembler.build(dirty_edits=dirty_edits, patch_fn=patch_sheet_worker)
```

---

## Usage Examples

### 1. Supported Data Types & Formulas

`SheetEditPlan.from_dict()` accepts cell coordinates as A1-style strings (`"B5"`, `"$AA$10"`) or 1-based `(row, col)` tuples (`(5, 2)`):

```python
from datetime import date, datetime, time
from sheet_patcher import SheetEditPlan

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

By default, writing to an existing cell keeps its existing `s` style index. To apply a known style index from the template or strip formatting completely, construct `CellEdit` / `RowEditBatch` objects directly or pass `styles` to `from_dict`:

```python
from sheet_patcher import CellEdit, RowEditBatch, SheetEditPlan

# Option A: Pass a style ID map alongside values
plan = SheetEditPlan.from_dict(
    cell_map={"A10": 9999.95, "B10": "Total"},
    styles={"A10": "14", "B10": "15"},  # Style indices from xl/styles.xml
)

# Option B: Explicitly clear styling on specific cells
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

### 3. Tuning Multiprocessing & Buffer Sizes

Control worker concurrency, I/O buffer size, and formula recalculation behavior on `PackageAssembler`:

```python
from package_pipeline import PackageAssembler
from sheet_patcher import patch_sheet_worker

assembler = PackageAssembler(
    src_path="large_workbook.xlsx",
    dst_path="large_workbook_patched.xlsx",
    max_workers=4,               # Defaults to os.process_cpu_count()
    chunk_size=2 * 1024 * 1024,  # 2 MB zip stream copy buffer
)

assembler.build(
    dirty_edits=dirty_edits,
    patch_fn=patch_sheet_worker,
    force_recalc=True,  # Drops calcChain.xml and sets fullCalcOnLoad="1"
)
```

---

## Architecture Overview

| Module | Class / Function | Responsibility |
| :--- | :--- | :--- |
| `package_pipeline.py` | `PackageReader` | Parses `xl/workbook.xml` and `xl/_rels/workbook.xml.rels` to map sheet names to internal zip paths (`xl/worksheets/sheetN.xml`). |
| `package_pipeline.py` | `PackageAssembler` | Dispatches dirty sheets to `ProcessPoolExecutor`, streams untouched zip entries in chunks, injects patched XMLs, and cleans `calcChain.xml`. |
| `sheet_patcher.py` | `SheetEditPlan` | Pre-sorts cell edits into ascending `(row, col)` order so the XML streamer can merge-join in a single linear pass. |
| `sheet_patcher.py` | `SheetStreamPatcher` | Streams `<worksheet>` via `iterparse` + `xmlfile`, inserts/patches `<row>` and `<c>` nodes, and prunes processed DOM nodes immediately. |
| `sheet_patcher.py` | `patch_sheet_worker` | Top-level picklable entrypoint executed inside each worker process. |

---

## Current MVP Limitations

* **Cell Value & Formula Updates Only:** Does not insert/delete rows (shifting formulas) or auto-resize Excel Tables (`xl/tables/tableN.xml`).
* **Date Cells Require Existing Date Formatting:** `date` and `datetime` values serialize to standard Excel serial floats; the target cell must already have a date number format style (`s`) applied in the template to render as a date rather than a number.
* **No New Style Creation:** Can inherit or assign existing `s` style indices from the workbook, but does not author new entries in `xl/styles.xml`.