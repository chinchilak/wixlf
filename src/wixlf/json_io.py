"""JSON record tables and worksheet value reading."""

import json
import math
import zipfile
from collections.abc import Iterator
from pathlib import Path

from lxml import etree

from .constants import (
    NS_MAIN,
    TAG_CELL,
    TAG_INLINE_STR,
    TAG_ROW,
    TAG_TEXT,
    TAG_VALUE,
    TAG_WORKSHEET,
    CellScalar,
)
from .models import Formula
from .reader import read_sheet_infos
from .utils.coords import check_cell_bounds, parse_coord
from .utils.xml_text import parse_xml, reject_doctype

type JSONValue = str | int | float | bool | None | list[JSONValue] | dict[str, JSONValue]
type JSONRecords = list[dict[str, JSONValue]]


def json_to_cells(
    records: JSONRecords | str,
    *,
    start_row: int = 1,
    start_col: int = 1,
    columns: list[str] | None = None,
    include_header: bool = True,
    parse_formulas: bool = False,
) -> dict[tuple[int, int], CellScalar]:
    """Convert records or a JSON array string into a 1-based cell mapping.

    Keys are collected in first-seen order unless columns are supplied. Missing
    fields become blank cells; nested objects and arrays become JSON text.
    """
    check_cell_bounds(start_row, start_col)
    if isinstance(records, str):
        records = json.loads(records)
    if not isinstance(records, list) or any(
        not isinstance(record, dict)
        or any(not isinstance(key, str) for key in record)
        for record in records
    ):
        raise ValueError("records must be a JSON array of objects with string keys.")
    if columns is None:
        columns = list(dict.fromkeys(key for record in records for key in record))
    elif any(not isinstance(column, str) for column in columns):
        raise ValueError("columns must contain only strings.")
    if len(set(columns)) != len(columns):
        raise ValueError("columns must not contain duplicate names.")
    if not columns:
        return {}
    row_count = len(records) + int(include_header)
    if row_count:
        check_cell_bounds(start_row + row_count - 1, start_col + len(columns) - 1)

    cells: dict[tuple[int, int], CellScalar] = {}
    if include_header:
        cells.update(
            ((start_row, start_col + offset), name)
            for offset, name in enumerate(columns)
        )
    for row_offset, record in enumerate(records, start=int(include_header)):
        for col_offset, name in enumerate(columns):
            value = record.get(name)
            if isinstance(value, (dict, list)):
                value = json.dumps(value, ensure_ascii=False, allow_nan=False)
            elif isinstance(value, float) and not math.isfinite(value):
                raise ValueError("JSON numbers must be finite.")
            elif value is not None and not isinstance(value, (str, bool, int, float)):
                raise TypeError(f"Unsupported JSON cell value for column {name!r}.")
            cell_value = (
                Formula(value)
                if parse_formulas and isinstance(value, str) and value.startswith("=")
                else value
            )
            cells[(start_row + row_offset, start_col + col_offset)] = cell_value
    return cells


def _text_content(element: etree._Element) -> str:
    # Only ordinary text and rich-text runs, not phonetic annotations (rPh).
    return "".join(
        text.text or ""
        for text in element.findall(TAG_TEXT)
        + element.findall(f"{{{NS_MAIN}}}r/{TAG_TEXT}")
    )


def _cell_value(cell: etree._Element, shared_strings: list[str]) -> JSONValue:
    cell_type = cell.get("t")
    if cell_type == "inlineStr":
        inline = cell.find(TAG_INLINE_STR)
        return _text_content(inline) if inline is not None else ""
    value = cell.find(TAG_VALUE)
    if value is None or value.text is None:
        return None
    raw = value.text
    if cell_type == "s":
        index = int(raw)
        if not 0 <= index < len(shared_strings):
            raise ValueError(f"Invalid shared-string index: {raw!r}")
        return shared_strings[index]
    if cell_type == "b":
        if raw not in ("0", "1"):
            raise ValueError(f"Invalid boolean cell value: {raw!r}")
        return raw == "1"
    if cell_type in ("str", "e", "d"):
        return raw
    try:
        return int(raw)
    except ValueError:
        number = float(raw)
        if not math.isfinite(number):
            raise ValueError(f"Non-finite worksheet number: {raw!r}")
        return number


def _worksheet_rows(
    archive: zipfile.ZipFile, path: str, shared_strings: list[str]
) -> Iterator[tuple[int, dict[int, JSONValue]]]:
    with archive.open(path) as stream:
        context = etree.iterparse(
            stream,
            events=("start", "end"),
            tag=(TAG_WORKSHEET, TAG_ROW),
            resolve_entities=False,
            no_network=True,
            load_dtd=False,
        )
        row_index = 0
        for event, element in context:
            if event == "start":
                if element.tag == TAG_WORKSHEET:
                    reject_doctype(element)
                continue
            if element.tag != TAG_ROW:
                continue
            row_index = int(element.get("r", str(row_index + 1)))
            values: dict[int, JSONValue] = {}
            col_index = 0
            for cell in element.findall(TAG_CELL):
                reference = cell.get("r")
                col_index = parse_coord(reference)[1] if reference else col_index + 1
                check_cell_bounds(row_index, col_index)
                values[col_index] = _cell_value(cell, shared_strings)
            yield row_index, values
            element.clear()
            parent = element.getparent()
            while element.getprevious() is not None:
                del parent[0]


def excel_to_json(
    xlsx_path: str | Path,
    sheet_name: str,
    *,
    header_row: int = 1,
    start_col: int = 1,
    as_json_str: bool = False,
) -> JSONRecords | str:
    """Read a header-based table as records or an indented JSON string.

    Blank headers are ignored; duplicate headers raise ValueError. Empty rows
    are skipped. Formulas use cached values (None when absent), dates are not
    style-decoded, and nested JSON text is not automatically deserialized.
    """
    check_cell_bounds(header_row, start_col)
    records: JSONRecords = []
    with zipfile.ZipFile(xlsx_path) as archive:
        sheets = read_sheet_infos(archive)
        if sheet_name not in sheets:
            raise KeyError(f"Worksheet {sheet_name!r} not found in {xlsx_path}")
        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in archive.namelist():
            root = parse_xml(archive.read("xl/sharedStrings.xml"))
            shared_strings = [
                _text_content(item) for item in root.findall(f"{{{NS_MAIN}}}si")
            ]
        headers: dict[int, str] = {}
        for row_index, values in _worksheet_rows(
            archive, sheets[sheet_name].zip_path, shared_strings
        ):
            if row_index < header_row:
                continue
            if row_index == header_row:
                headers = {
                    col: str(value)
                    for col, value in sorted(values.items())
                    if col >= start_col and value is not None and value != ""
                }
                if len(set(headers.values())) != len(headers):
                    raise ValueError("Worksheet headers must be unique.")
            elif headers:
                record = {name: values.get(col) for col, name in headers.items()}
                if any(value is not None for value in record.values()):
                    records.append(record)
    return json.dumps(records, indent=2, ensure_ascii=False, allow_nan=False) if as_json_str else records
