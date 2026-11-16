import zipfile
from pathlib import Path

from lxml import etree

from .constants import NS_MAIN, NS_REL_DOC, NS_REL_PKG
from .models import SheetInfo
from .utils.zip_io import resolve_zip_path


class PackageReader:
    """
    Reads small OOXML metadata files from an existing .xlsx/.xlsm archive
    to resolve human-readable sheet names to internal zip paths.
    """

    def __init__(self, src_path: str | Path) -> None:
        self.src_path = Path(src_path)

    def get_sheets(self) -> dict[str, SheetInfo]:
        """Returns a mapping of {sheet_name: SheetInfo}."""
        with zipfile.ZipFile(self.src_path, "r") as zin:
            wb_rels = self._parse_relationships(zin, "xl/_rels/workbook.xml.rels")
            wb_root = etree.fromstring(zin.read("xl/workbook.xml"))

        sheets: dict[str, SheetInfo] = {}
        for sheet_elem in wb_root.findall(f".//{{{NS_MAIN}}}sheet"):
            name = sheet_elem.get("name")
            sheet_id = sheet_elem.get("sheetId", "")
            rel_id = sheet_elem.get(f"{{{NS_REL_DOC}}}id")

            if not name or not rel_id or rel_id not in wb_rels:
                continue

            target = wb_rels[rel_id]
            zip_path = resolve_zip_path("xl/workbook.xml", target)
            sheets[name] = SheetInfo(
                name=name,
                sheet_id=sheet_id,
                rel_id=rel_id,
                zip_path=zip_path,
            )

        return sheets

    @staticmethod
    def _parse_relationships(zin: zipfile.ZipFile, rels_path: str) -> dict[str, str]:
        """Maps rId -> Target path."""
        rels_root = etree.fromstring(zin.read(rels_path))
        return {
            rel_id: target
            for rel in rels_root.findall(f"{{{NS_REL_PKG}}}Relationship")
            if (rel_id := rel.get("Id")) and (target := rel.get("Target"))
        }
