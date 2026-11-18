import zipfile
from collections.abc import Sequence
from pathlib import Path

from .constants import NS_MAIN, NS_REL_DOC
from .models import SheetInfo
from .utils.xml_text import parse_xml
from .utils.zip_io import Relationship, read_relationships, resolve_zip_path

WORKBOOK_PATH = "xl/workbook.xml"
WORKBOOK_RELS_PATH = "xl/_rels/workbook.xml.rels"


def read_sheet_infos(
    zin: zipfile.ZipFile,
    wb_rels: Sequence[Relationship] | None = None,
) -> dict[str, SheetInfo]:
    """Returns an insertion-ordered {sheet_name: SheetInfo}; pass wb_rels to reuse an already-parsed rels list."""
    if wb_rels is None:
        wb_rels = read_relationships(zin, WORKBOOK_RELS_PATH)
    target_by_id = {rel.id: rel.target for rel in wb_rels}
    wb_root = parse_xml(zin.read(WORKBOOK_PATH))

    sheets: dict[str, SheetInfo] = {}
    for sheet_elem in wb_root.findall(f".//{{{NS_MAIN}}}sheet"):
        name = sheet_elem.get("name")
        rel_id = sheet_elem.get(f"{{{NS_REL_DOC}}}id")
        if not name or not rel_id or rel_id not in target_by_id:
            continue
        sheets[name] = SheetInfo(
            name=name,
            sheet_id=sheet_elem.get("sheetId", ""),
            rel_id=rel_id,
            zip_path=resolve_zip_path(WORKBOOK_PATH, target_by_id[rel_id]),
            state=sheet_elem.get("state", "visible"),
        )
    return sheets


class PackageReader:
    """
    Reads small OOXML metadata files from an existing .xlsx/.xlsm archive
    to resolve human-readable sheet names to internal zip paths in tab order.
    """

    def __init__(self, src_path: str | Path) -> None:
        self.src_path = Path(src_path)

    def get_sheets(self) -> dict[str, SheetInfo]:
        """Returns an insertion-ordered mapping of {sheet_name: SheetInfo}."""
        with zipfile.ZipFile(self.src_path, "r") as zin:
            return read_sheet_infos(zin)
