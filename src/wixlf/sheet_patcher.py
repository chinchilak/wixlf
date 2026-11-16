import math
import zipfile
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from datetime import date, datetime, time
from typing import Final

from lxml import etree

NS_MAIN: Final[str] = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_XML: Final[str] = "http://www.w3.org/XML/1998/namespace"

TAG_WORKSHEET: Final[str] = f"{{{NS_MAIN}}}worksheet"
TAG_SHEET_DATA: Final[str] = f"{{{NS_MAIN}}}sheetData"
TAG_DIMENSION: Final[str] = f"{{{NS_MAIN}}}dimension"
TAG_ROW: Final[str] = f"{{{NS_MAIN}}}row"
TAG_CELL: Final[str] = f"{{{NS_MAIN}}}c"
TAG_VALUE: Final[str] = f"{{{NS_MAIN}}}v"
TAG_FORMULA: Final[str] = f"{{{NS_MAIN}}}f"
TAG_INLINE_STR: Final[str] = f"{{{NS_MAIN}}}is"
TAG_TEXT: Final[str] = f"{{{NS_MAIN}}}t"

type CellScalar = str | int | float | bool | date | datetime | time | None

# Translation table mapping illegal XML 1.0 characters to None (fast C-level str.translate)
_ILLEGAL_XML_TRANS: Final[dict[int, None]] = dict.fromkeys(
    [
        *range(0x00, 0x09),      # NUL..BS (keeps 0x09 TAB, 0x0A LF)
        0x0B,                    # VT
        0x0C,                    # FF (keeps 0x0D CR)
        *range(0x0E, 0x20),      # SO..US
        *range(0xD800, 0xE000),  # UTF-16 surrogates
        0xFFFE,
        0xFFFF,
    ],
    None,
)

# Excel 1900 epoch (with Excel's intentional 1900 leap-year bug offset)
_EXCEL_EPOCH: Final[datetime] = datetime(1899, 12, 30)


def sanitize_xml_text(text: str) -> str:
    """Strips characters forbidden in XML 1.0 using C-accelerated str.translate."""
    return text.translate(_ILLEGAL_XML_TRANS)


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
    cleaned = cleaned.removeprefix("$")

    split_idx = 0
    length = len(cleaned)
    while split_idx < length and cleaned[split_idx].isascii() and cleaned[split_idx].isalpha():
        split_idx += 1

    col_letters = cleaned[:split_idx]
    row_part = cleaned[split_idx:]
    row_part = row_part.removeprefix("$")

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


@dataclass(frozen=True, slots=True)
class CellEdit:
    """Represents a single cell mutation."""

    row: int
    col: int
    value: CellScalar
    style_id: str | None = None  # Optional explicit 's' attribute override
    clear_style: bool = False


@dataclass(frozen=True, slots=True)
class RowEditBatch:
    """Sorted cell mutations belonging to a single row."""

    row: int
    cells: tuple[CellEdit, ...]  # Strictly sorted by col ascending


@dataclass(frozen=True, slots=True)
class SheetEditPlan:
    """
    Pre-sorted, immutable, picklable edit plan for a single worksheet.
    Passed across ProcessPoolExecutor boundaries.
    """

    rows: tuple[RowEditBatch, ...]  # Strictly sorted by row ascending

    @classmethod
    def from_dict(
        cls,
        cell_map: Mapping[str | tuple[int, int], CellScalar],
        styles: Mapping[str | tuple[int, int], str] | None = None,
    ) -> "SheetEditPlan":
        """
        Compiles an unsorted dict like {'B5': 42, (1, 1): 'Title'} into a
        strictly sorted SheetEditPlan ready for single-pass streaming.
        """
        by_row: dict[int, dict[int, CellEdit]] = {}
        style_lookup: dict[tuple[int, int], str] = {}

        if styles:
            for key, s_id in styles.items():
                r, c = parse_coord(key) if isinstance(key, str) else key
                style_lookup[(r, c)] = str(s_id)

        for key, val in cell_map.items():
            r, c = parse_coord(key) if isinstance(key, str) else key
            if r < 1 or c < 1:
                raise ValueError(f"Row and column must be >= 1, got ({r}, {c})")
            by_row.setdefault(r, {})[c] = CellEdit(
                row=r,
                col=c,
                value=val,
                style_id=style_lookup.get((r, c)),
            )

        sorted_batches = tuple(
            RowEditBatch(
                row=r,
                cells=tuple(by_row[r][c] for c in sorted(by_row[r])),
            )
            for r in sorted(by_row)
        )
        return cls(rows=sorted_batches)


class SheetStreamPatcher:
    """
    Streams an existing worksheet XML from a zip archive and writes a patched
    version to disk using constant O(1) memory.
    """

    def __init__(self, drop_dimension_tag: bool = True) -> None:
        # Dropping <dimension ref="A1:D10"/> prevents stale bounds if new rows/cols
        # extend beyond the original range; Excel recalculates it cleanly on open.
        self.drop_dimension_tag = drop_dimension_tag

    def patch_from_zip(
        self,
        src_zip_path: str,
        sheet_zip_path: str,
        plan: SheetEditPlan,
        out_xml_path: str,
    ) -> None:
        """Entrypoint compatible with PackageAssembler's PatchWorkerFn signature."""
        with (
            zipfile.ZipFile(src_zip_path, "r") as zin,
            zin.open(sheet_zip_path, "r") as xml_in,
            etree.xmlfile(out_xml_path, encoding="utf-8") as xf,
        ):
            xf.write_declaration(standalone=True)
            self._stream_transform(xml_in, xf, plan.rows)

    def _stream_transform(
        self,
        xml_in: zipfile.ZipExtFile,
        xf: etree.xmlfile,
        row_batches: Sequence[RowEditBatch],
    ) -> None:
        context = etree.iterparse(
            xml_in,
            events=("start", "end"),
            recover=False,
            huge_tree=True,
        )

        # Peek at root <worksheet> start event to preserve exact namespaces/attrs
        event, root_elem = next(context)
        if event != "start" or root_elem.tag != TAG_WORKSHEET:
            raise ValueError(f"Expected <worksheet> root, got {root_elem.tag}")

        ws_attrib = dict(root_elem.attrib)
        ws_nsmap = dict(root_elem.nsmap)

        with xf.element(TAG_WORKSHEET, attrib=ws_attrib, nsmap=ws_nsmap):
            in_sheet_data = False
            sheet_data_written = False
            batch_idx = 0
            total_batches = len(row_batches)

            for event, elem in context:
                tag = elem.tag

                # Track entry/exit of <sheetData>
                if event == "start" and tag == TAG_SHEET_DATA:
                    in_sheet_data = True
                    sheet_data_written = True
                    sd_attrib = dict(elem.attrib)
                    sd_ctx = xf.element(TAG_SHEET_DATA, attrib=sd_attrib)
                    sd_ctx.__enter__()
                    continue

                if event == "end" and tag == TAG_SHEET_DATA:
                    # Flush any remaining new rows that belong after the last existing row
                    while batch_idx < total_batches:
                        self._write_new_row(xf, row_batches[batch_idx])
                        batch_idx += 1

                    sd_ctx.__exit__(None, None, None)
                    in_sheet_data = False
                    self._prune_element(elem)
                    continue

                # Inside <sheetData>: merge-join sorted edits against <row> stream
                if in_sheet_data:
                    if event == "end" and tag == TAG_ROW:
                        row_num = int(elem.get("r", "0"))

                        # Insert any brand-new rows strictly preceding this existing row
                        while (
                            batch_idx < total_batches
                            and row_batches[batch_idx].row < row_num
                        ):
                            self._write_new_row(xf, row_batches[batch_idx])
                            batch_idx += 1

                        # Patch existing row if it matches the current edit batch
                        if (
                            batch_idx < total_batches
                            and row_batches[batch_idx].row == row_num
                        ):
                            self._patch_existing_row(elem, row_batches[batch_idx])
                            batch_idx += 1

                        xf.write(elem)
                        self._prune_element(elem)
                    continue

                # Outside <sheetData>: pass top-level worksheet children through verbatim
                if event == "end" and elem.getparent() is root_elem:
                    if self.drop_dimension_tag and tag == TAG_DIMENSION:
                        self._prune_element(elem)
                        continue

                    xf.write(elem)
                    self._prune_element(elem)

            # Edge case: source worksheet had no <sheetData> tag at all
            if not sheet_data_written and total_batches > 0:
                with xf.element(TAG_SHEET_DATA):
                    for batch in row_batches:
                        self._write_new_row(xf, batch)

        root_elem.clear()

    def _patch_existing_row(
        self, row_elem: etree._Element, batch: RowEditBatch
    ) -> None:
        """
        Merges sorted CellEdits into an existing <row> element in linear time,
        preserving existing cell styles ('s') unless overridden.
        """
        existing_cells = list(row_elem)
        new_children: list[etree._Element] = []

        c_idx = 0
        e_idx = 0
        num_existing = len(existing_cells)
        num_edits = len(batch.cells)

        while c_idx < num_existing and e_idx < num_edits:
            cell_elem = existing_cells[c_idx]
            edit = batch.cells[e_idx]
            _, existing_col = parse_coord(cell_elem.get("r", "A1"))

            if existing_col < edit.col:
                new_children.append(cell_elem)
                c_idx += 1
            elif existing_col > edit.col:
                new_children.append(self._build_cell_element(edit, inherited_style=None))
                e_idx += 1
            else:
                # Overwrite existing cell; inherit style unless explicitly changed
                inherited_style = cell_elem.get("s")
                new_children.append(
                    self._build_cell_element(edit, inherited_style=inherited_style)
                )
                c_idx += 1
                e_idx += 1

        while c_idx < num_existing:
            new_children.append(existing_cells[c_idx])
            c_idx += 1

        while e_idx < num_edits:
            new_children.append(
                self._build_cell_element(batch.cells[e_idx], inherited_style=None)
            )
            e_idx += 1

        # Replace row children in-place while keeping <row> attributes (ht, customHeight, s, spans)
        row_elem[:] = new_children
        # Clear cached spans attribute if new columns may have widened the row
        row_elem.attrib.pop("spans", None)

    def _write_new_row(self, xf: etree.xmlfile, batch: RowEditBatch) -> None:
        """Streams a brand-new <row> element directly to xmlfile."""
        row_elem = etree.Element(TAG_ROW, r=str(batch.row))
        for edit in batch.cells:
            row_elem.append(self._build_cell_element(edit, inherited_style=None))
        xf.write(row_elem)

    def _build_cell_element(
        self, edit: CellEdit, inherited_style: str | None
    ) -> etree._Element:
        """Constructs a compliant OOXML <c> element for any scalar or formula."""
        coord = format_coord(edit.row, edit.col)
        c_elem = etree.Element(TAG_CELL, r=coord)

        # Resolve style index ('s')
        if not edit.clear_style:
            final_style = edit.style_id if edit.style_id is not None else inherited_style
            if final_style is not None:
                c_elem.set("s", final_style)

        val = edit.value

        # 1. Empty / cleared cell (keeps style if present, omits type and value)
        if val is None:
            return c_elem

        # 2. Boolean (must be checked before int because bool is a subclass of int)
        if isinstance(val, bool):
            c_elem.set("t", "b")
            v_elem = etree.SubElement(c_elem, TAG_VALUE)
            v_elem.text = "1" if val else "0"
            return c_elem

        # 3. Numeric (int / float)
        if isinstance(val, (int, float)):
            if isinstance(val, float) and (math.isnan(val) or math.isinf(val)):
                c_elem.set("t", "e")
                v_elem = etree.SubElement(c_elem, TAG_VALUE)
                v_elem.text = "#NUM!"
            else:
                # Default OOXML cell type is numeric ('n' can be omitted)
                v_elem = etree.SubElement(c_elem, TAG_VALUE)
                v_elem.text = repr(val) if isinstance(val, float) else str(val)
            return c_elem

        # 4. Date / Datetime / Time -> Excel serial number
        if isinstance(val, (datetime, date, time)):
            c_elem.set("t", "n")
            v_elem = etree.SubElement(c_elem, TAG_VALUE)
            v_elem.text = repr(self._to_excel_serial(val))
            return c_elem

        # 5. Formula vs. Inline String
        if isinstance(val, str):
            if val.startswith("="):
                f_elem = etree.SubElement(c_elem, TAG_FORMULA)
                f_elem.text = val[1:]  # OOXML stores formulas without leading '='
                # Omit <v> so Excel computes the fresh result on open
            else:
                clean_text = sanitize_xml_text(val)
                c_elem.set("t", "inlineStr")
                is_elem = etree.SubElement(c_elem, TAG_INLINE_STR)
                t_elem = etree.SubElement(is_elem, TAG_TEXT)
                if clean_text != clean_text.strip() or "\n" in clean_text:
                    t_elem.set(f"{{{NS_XML}}}space", "preserve")
                t_elem.text = clean_text
            return c_elem

        raise TypeError(f"Unsupported cell value type: {type(val).__name__}")

    @staticmethod
    def _to_excel_serial(dt_val: datetime | date | time) -> float:
        """Converts Python date/datetime/time into Excel 1900 serial float."""
        if isinstance(dt_val, datetime):
            naive = dt_val.replace(tzinfo=None)
            delta = naive - _EXCEL_EPOCH
            serial = delta.days + (delta.seconds + delta.microseconds / 1_000_000) / 86400.0
        elif isinstance(dt_val, date):
            naive = datetime(dt_val.year, dt_val.month, dt_val.day)
            serial = float((naive - _EXCEL_EPOCH).days)
        else:  # time
            serial = (
                dt_val.hour * 3600
                + dt_val.minute * 60
                + dt_val.second
                + dt_val.microsecond / 1_000_000
            ) / 86400.0

        # Adjust for Excel's fictional Feb 29, 1900 (serial 60)
        if not isinstance(dt_val, time) and serial <= 60:
            serial -= 1.0
        return serial

    @staticmethod
    def _prune_element(elem: etree._Element) -> None:
        """Frees element memory and removes all already-processed preceding siblings."""
        elem.clear()
        parent = elem.getparent()
        if parent is not None:
            while elem.getprevious() is not None:
                del parent[0]


# Top-level picklable worker function for PackageAssembler
def patch_sheet_worker(
    src_zip_path: str,
    sheet_zip_path: str,
    plan: SheetEditPlan,
    out_xml_path: str,
) -> None:
    patcher = SheetStreamPatcher()
    patcher.patch_from_zip(src_zip_path, sheet_zip_path, plan, out_xml_path)
