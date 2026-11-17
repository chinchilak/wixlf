import math
import zipfile
from collections.abc import Sequence
from datetime import date, datetime, time

from lxml import etree

from .constants import (
    MAX_EXCEL_COLS,
    NS_MAIN,
    NS_REL_DOC,
    NS_XML,
    TAG_AUTO_FILTER,
    TAG_CELL,
    TAG_COL,
    TAG_COLS,
    TAG_COND_FMT,
    TAG_DATA_VAL,
    TAG_DATA_VALS,
    TAG_DIMENSION,
    TAG_FORMULA,
    TAG_INLINE_STR,
    TAG_MERGE_CELL,
    TAG_MERGE_CELLS,
    TAG_ROW,
    TAG_SHEET_DATA,
    TAG_TEXT,
    TAG_VALUE,
    TAG_WORKSHEET,
)
from .models import (
    CellEdit,
    RowEditBatch,
    SheetEditPlan,
    SheetStructureMutation,
    WorksheetTaskPayload,
)
from .utils.coords import format_coord, parse_coord
from .utils.shifting import (
    AxisShiftOp,
    rewrite_formula,
    shift_index,
    shift_range_str,
    shift_sqref,
)
from .utils.xml_text import sanitize_xml_text, to_excel_serial


class SheetStreamPatcher:
    """
    Streams an existing worksheet XML from a zip archive (or generates a new
    worksheet XML from scratch) and writes to disk using constant O(1) memory.
    Supports row/column insertions & deletions with automatic formula & range shifting.
    """

    def __init__(self, drop_dimension_tag: bool = True) -> None:
        # Dropping <dimension ref="A1:D10"/> prevents stale bounds if rows/cols
        # shift or extend beyond the original range; Excel recalculates it on open.
        self.drop_dimension_tag = drop_dimension_tag

    def patch_from_zip(
        self,
        src_zip_path: str,
        sheet_zip_path: str,
        payload: SheetEditPlan | WorksheetTaskPayload,
        out_xml_path: str,
    ) -> None:
        """Streams and patches an existing worksheet entry inside src_zip_path."""
        if isinstance(payload, WorksheetTaskPayload):
            plan = payload.plan
            mutation = payload.mutation
        else:
            plan = payload
            mutation = None

        with (
            zipfile.ZipFile(src_zip_path, "r") as zin,
            zin.open(sheet_zip_path, "r") as xml_in,
            etree.xmlfile(out_xml_path, encoding="utf-8") as xf,
        ):
            xf.write_declaration(standalone=True)
            self._stream_transform(xml_in, xf, plan.rows, mutation)

    def write_new_sheet(
        self,
        plan: SheetEditPlan,
        out_xml_path: str,
    ) -> None:
        """Streams a brand-new compliant worksheet XML file to out_xml_path."""
        nsmap = {None: NS_MAIN, "r": NS_REL_DOC}
        with etree.xmlfile(out_xml_path, encoding="utf-8") as xf:
            xf.write_declaration(standalone=True)
            with (
                xf.element(TAG_WORKSHEET, nsmap=nsmap),
                xf.element(TAG_SHEET_DATA),
            ):
                for batch in plan.rows:
                    self._write_new_row(xf, batch, mutation=None)

    def _stream_transform(
        self,
        xml_in: zipfile.ZipExtFile,
        xf: etree.xmlfile,
        row_batches: Sequence[RowEditBatch],
        mutation: SheetStructureMutation | None,
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

        row_ops = mutation.row_ops if mutation else ()
        col_ops = mutation.col_ops if mutation else ()
        has_local_shifts = bool(row_ops or col_ops)
        has_any_shifts = mutation.has_any_shifts if mutation else False

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
                        self._write_new_row(xf, row_batches[batch_idx], mutation)
                        batch_idx += 1

                    sd_ctx.__exit__(None, None, None)
                    in_sheet_data = False
                    self._prune_element(elem)
                    continue

                # Inside <sheetData>: shift coordinates & merge-join sorted cell edits
                if in_sheet_data:
                    if event == "end" and tag == TAG_ROW:
                        raw_row_num = int(elem.get("r", "0"))
                        target_row = (
                            shift_index(raw_row_num, row_ops)
                            if row_ops
                            else raw_row_num
                        )

                        # Row fell inside a deleted row span -> drop entire row
                        if target_row is None:
                            self._prune_element(elem)
                            continue

                        # Shift row index, cell columns, and formulas in-place on <row>
                        if has_any_shifts:
                            self._shift_row_in_place(elem, target_row, mutation)

                        # Insert any brand-new rows strictly preceding this row
                        while (
                            batch_idx < total_batches
                            and row_batches[batch_idx].row < target_row
                        ):
                            self._write_new_row(xf, row_batches[batch_idx], mutation)
                            batch_idx += 1

                        # Patch existing row if it matches the current edit batch
                        if (
                            batch_idx < total_batches
                            and row_batches[batch_idx].row == target_row
                        ):
                            self._patch_existing_row(
                                elem, row_batches[batch_idx], mutation
                            )
                            batch_idx += 1

                        # Write row only if it still has cells or non-trivial row attrs
                        if len(elem) > 0 or len(elem.attrib) > 1:
                            xf.write(elem)
                        self._prune_element(elem)
                    continue

                # Outside <sheetData>: pass/transform top-level worksheet children
                if event == "end" and elem.getparent() is root_elem:
                    if self.drop_dimension_tag and tag == TAG_DIMENSION:
                        self._prune_element(elem)
                        continue

                    if has_local_shifts and not self._transform_worksheet_metadata(elem, row_ops, col_ops):
                        self._prune_element(elem)
                        continue

                    xf.write(elem)
                    self._prune_element(elem)

            # Edge case: source worksheet had no <sheetData> tag at all
            if not sheet_data_written and total_batches > 0:
                with xf.element(TAG_SHEET_DATA):
                    for batch in row_batches:
                        self._write_new_row(xf, batch, mutation)

        root_elem.clear()

    def _shift_row_in_place(
        self,
        row_elem: etree._Element,
        target_row: int,
        mutation: SheetStructureMutation | None,
    ) -> None:
        """Shifts <row r="...">, child <c r="..."> coordinates, and <f> formulas."""
        row_elem.set("r", str(target_row))
        col_ops = mutation.col_ops if mutation else ()
        if col_ops:
            row_elem.attrib.pop("spans", None)

        kept_cells: list[etree._Element] = []
        for cell_elem in list(row_elem):
            _, orig_col = parse_coord(cell_elem.get("r", "A1"))
            new_col = shift_index(orig_col, col_ops) if col_ops else orig_col
            if new_col is None:
                continue  # Cell fell in deleted column span

            cell_elem.set("r", format_coord(target_row, new_col))

            if mutation and mutation.has_any_shifts:
                f_elem = cell_elem.find(TAG_FORMULA)
                if f_elem is not None and f_elem.text:
                    f_elem.text = self._rewrite_formula_text(f_elem.text, mutation)
                    # Drop cached formula result <v> so Excel computes fresh value
                    v_elem = cell_elem.find(TAG_VALUE)
                    if v_elem is not None:
                        cell_elem.remove(v_elem)

            kept_cells.append(cell_elem)

        row_elem[:] = kept_cells

    def _patch_existing_row(
        self,
        row_elem: etree._Element,
        batch: RowEditBatch,
        mutation: SheetStructureMutation | None,
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
                if not edit.drop_cell:
                    new_children.append(
                        self._build_cell_element(
                            edit, inherited_style=None, mutation=mutation
                        )
                    )
                e_idx += 1
            else:
                # Existing cell matches target column
                if not edit.drop_cell:
                    inherited_style = cell_elem.get("s")
                    new_children.append(
                        self._build_cell_element(
                            edit, inherited_style=inherited_style, mutation=mutation
                        )
                    )
                c_idx += 1
                e_idx += 1

        while c_idx < num_existing:
            new_children.append(existing_cells[c_idx])
            c_idx += 1

        while e_idx < num_edits:
            edit = batch.cells[e_idx]
            if not edit.drop_cell:
                new_children.append(
                    self._build_cell_element(edit, inherited_style=None, mutation=mutation)
                )
            e_idx += 1

        # Replace row children in-place while keeping <row> attributes (ht, customHeight, s)
        row_elem[:] = new_children
        # Clear cached spans attribute if new columns may have widened the row
        row_elem.attrib.pop("spans", None)

    def _write_new_row(
        self,
        xf: etree.xmlfile,
        batch: RowEditBatch,
        mutation: SheetStructureMutation | None,
    ) -> None:
        """Streams a brand-new <row> element directly to xmlfile."""
        cells_to_write = [e for e in batch.cells if not e.drop_cell]
        if not cells_to_write:
            return
        row_elem = etree.Element(TAG_ROW, r=str(batch.row))
        for edit in cells_to_write:
            row_elem.append(
                self._build_cell_element(edit, inherited_style=None, mutation=mutation)
            )
        xf.write(row_elem)

    def _build_cell_element(
        self,
        edit: CellEdit,
        inherited_style: str | None,
        mutation: SheetStructureMutation | None = None,
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
                v_elem = etree.SubElement(c_elem, TAG_VALUE)
                v_elem.text = repr(val) if isinstance(val, float) else str(val)
            return c_elem

        # 4. Date / Datetime / Time -> Excel serial number
        if isinstance(val, (datetime, date, time)):
            c_elem.set("t", "n")
            v_elem = etree.SubElement(c_elem, TAG_VALUE)
            v_elem.text = repr(to_excel_serial(val))
            return c_elem

        # 5. Formula vs. Inline String
        if isinstance(val, str):
            if val.startswith("="):
                f_elem = etree.SubElement(c_elem, TAG_FORMULA)
                formula_body = val[1:]  # OOXML stores formulas without leading '='
                if mutation and mutation.has_any_shifts:
                    formula_body = self._rewrite_formula_text(formula_body, mutation)
                f_elem.text = formula_body
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
    def _rewrite_formula_text(
        formula_body: str, mutation: SheetStructureMutation
    ) -> str:
        shifts_map = mutation.all_shifts or {
            mutation.sheet_name: (mutation.row_ops, mutation.col_ops)
        }
        return rewrite_formula(
            formula_body,
            current_sheet=mutation.sheet_name,
            shifts_by_sheet=shifts_map,
            deleted_sheets=mutation.deleted_sheets,
        )

    @staticmethod
    def _transform_worksheet_metadata(
        elem: etree._Element,
        row_ops: Sequence[AxisShiftOp],
        col_ops: Sequence[AxisShiftOp],
    ) -> bool:
        """
        Shifts worksheet-level coordinate metadata (<cols>, <mergeCells>, <autoFilter>,
        <conditionalFormatting>, <dataValidations>). Returns False if element becomes empty.
        """
        tag = elem.tag

        if tag == TAG_COLS and col_ops:
            new_cols: list[etree._Element] = []
            for col_el in list(elem.findall(TAG_COL)):
                c_min = int(col_el.get("min", "1"))
                c_max = int(col_el.get("max", str(c_min)))
                for op in col_ops:
                    if op.delta > 0:
                        if c_min >= op.start:
                            c_min += op.delta
                        if c_max >= op.start:
                            c_max += op.delta
                    elif op.delta < 0:
                        d_start = op.start
                        d_end = op.start + (-op.delta) - 1
                        if c_min >= d_start and c_max <= d_end:
                            c_min, c_max = 1, 0
                            break
                        before = max(0, min(c_min - 1, d_end) - d_start + 1)
                        up_to = max(0, min(c_max, d_end) - d_start + 1)
                        c_min -= before
                        c_max -= up_to
                if c_min <= c_max:
                    col_el.set("min", str(max(1, c_min)))
                    col_el.set("max", str(min(MAX_EXCEL_COLS, c_max)))
                    new_cols.append(col_el)
            elem[:] = new_cols
            return len(new_cols) > 0

        if tag == TAG_MERGE_CELLS:
            kept_merges: list[etree._Element] = []
            for mc in list(elem.findall(TAG_MERGE_CELL)):
                ref = mc.get("ref", "")
                if not ref:
                    continue
                new_ref = shift_range_str(ref, row_ops, col_ops)
                if new_ref is not None and ":" in new_ref:
                    mc.set("ref", new_ref)
                    kept_merges.append(mc)
            elem[:] = kept_merges
            if not kept_merges:
                return False
            elem.set("count", str(len(kept_merges)))
            return True

        if tag == TAG_AUTO_FILTER:
            ref = elem.get("ref", "")
            if ref:
                new_ref = shift_range_str(ref, row_ops, col_ops)
                if new_ref is None:
                    return False
                elem.set("ref", new_ref)
            return True

        if tag == TAG_COND_FMT:
            sqref = elem.get("sqref", "")
            if sqref:
                new_sqref = shift_sqref(sqref, row_ops, col_ops)
                if not new_sqref:
                    return False
                elem.set("sqref", new_sqref)
            return True

        if tag == TAG_DATA_VALS:
            kept_dvs: list[etree._Element] = []
            for dv in list(elem.findall(TAG_DATA_VAL)):
                sqref = dv.get("sqref", "")
                if sqref:
                    new_sqref = shift_sqref(sqref, row_ops, col_ops)
                    if not new_sqref:
                        continue
                    dv.set("sqref", new_sqref)
                kept_dvs.append(dv)
            elem[:] = kept_dvs
            if not kept_dvs:
                return False
            elem.set("count", str(len(kept_dvs)))
            return True

        return True

    @staticmethod
    def _prune_element(elem: etree._Element) -> None:
        """Frees element memory and removes all already-processed preceding siblings."""
        elem.clear()
        parent = elem.getparent()
        if parent is not None:
            while elem.getprevious() is not None:
                del parent[0]


# Top-level picklable worker functions for PackageAssembler
def patch_sheet_worker(
    src_zip_path: str,
    sheet_zip_path: str,
    payload: SheetEditPlan | WorksheetTaskPayload,
    out_xml_path: str,
) -> None:
    patcher = SheetStreamPatcher()
    patcher.patch_from_zip(src_zip_path, sheet_zip_path, payload, out_xml_path)


def create_sheet_worker(
    _src_zip_path: str,
    _sheet_zip_path: str,
    plan: SheetEditPlan,
    out_xml_path: str,
) -> None:
    patcher = SheetStreamPatcher()
    patcher.write_new_sheet(plan, out_xml_path)
