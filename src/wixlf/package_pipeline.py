import os
import posixpath
import shutil
import tempfile
import zipfile
from collections.abc import Callable, Mapping
from concurrent.futures import ProcessPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from lxml import etree

# OOXML Namespaces
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"

CALC_CHAIN_PATH = "xl/calcChain.xml"
CALC_CHAIN_CT = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"
)
CALC_CHAIN_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain"
)

# PEP 695 native type statement
# Signature: (src_zip_path, internal_sheet_path, sheet_edits, output_temp_xml_path) -> None
type PatchWorkerFn = Callable[[str, str, Any, str], None]


@dataclass(frozen=True, slots=True)
class SheetInfo:
    name: str
    sheet_id: str
    rel_id: str
    zip_path: str  # e.g., "xl/worksheets/sheet1.xml"


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
            zip_path = self._resolve_zip_path("xl/workbook.xml", target)
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

    @staticmethod
    def _resolve_zip_path(base_file: str, target: str) -> str:
        """Resolves relative or root-anchored OOXML targets to zip entry names."""
        if target.startswith("/"):
            return target.lstrip("/")
        base_dir = posixpath.dirname(base_file)
        return posixpath.normpath(posixpath.join(base_dir, target))


def _worker_entrypoint(
    patch_fn: PatchWorkerFn,
    src_path: str,
    sheet_zip_path: str,
    edits: Any,
    temp_dir: str,
) -> tuple[str, str]:
    """
    Top-level picklable wrapper executed in worker processes.
    Writes patched XML to a temp file and returns (sheet_zip_path, temp_file_path).
    """
    fd, tmp_xml_path = tempfile.mkstemp(
        prefix="sheet_patch_", suffix=".xml", dir=temp_dir
    )
    os.close(fd)
    try:
        patch_fn(src_path, sheet_zip_path, edits, tmp_xml_path)
        return sheet_zip_path, tmp_xml_path
    except Exception:
        if os.path.exists(tmp_xml_path):
            os.remove(tmp_xml_path)
        raise


class PackageAssembler:
    """
    Coordinates multiprocessing across dirty sheets and streams untouched + patched
    entries into the output .xlsx archive while cleanly stripping calcChain.xml.
    """

    # Avoid process spawn overhead when modifying a single tiny worksheet (< 512 KB)
    PARALLEL_SIZE_THRESHOLD_BYTES = 512 * 1024

    def __init__(
        self,
        src_path: str | Path,
        dst_path: str | Path,
        max_workers: int | None = None,
        chunk_size: int = 1024 * 1024,  # 1 MB streaming buffer
    ) -> None:
        self.src_path = Path(src_path)
        self.dst_path = Path(dst_path)
        self.max_workers = max_workers or os.process_cpu_count() or 1
        self.chunk_size = chunk_size

    def build(
        self,
        dirty_edits: Mapping[str, Any],
        patch_fn: PatchWorkerFn,
        force_recalc: bool = True,
    ) -> None:
        """
        dirty_edits: {sheet_zip_path: sorted_edit_plan}
        patch_fn: Top-level picklable function implementing SheetStreamPatcher
        force_recalc: Removes calcChain.xml and sets fullCalcOnLoad="1"
        """
        if not dirty_edits:
            shutil.copy2(self.src_path, self.dst_path)
            return

        with tempfile.TemporaryDirectory(prefix="xlsx_build_") as temp_dir:
            patched_files = self._run_sheet_workers(dirty_edits, patch_fn, temp_dir)
            self._assemble_zip(patched_files, strip_calc_chain=force_recalc)

    def _run_sheet_workers(
        self,
        dirty_edits: Mapping[str, Any],
        patch_fn: PatchWorkerFn,
        temp_dir: str,
    ) -> dict[str, str]:
        """Dispatches dirty sheets to ProcessPoolExecutor or runs inline if small."""
        with zipfile.ZipFile(self.src_path, "r") as zin:
            sheet_sizes = {
                path: zin.getinfo(path).file_size for path in dirty_edits
            }

        use_multiprocessing = len(dirty_edits) > 1 or any(
            sz >= self.PARALLEL_SIZE_THRESHOLD_BYTES for sz in sheet_sizes.values()
        )

        patched_map: dict[str, str] = {}
        src_str = str(self.src_path)

        if not use_multiprocessing:
            for sheet_path, edits in dirty_edits.items():
                path, tmp_file = _worker_entrypoint(
                    patch_fn, src_str, sheet_path, edits, temp_dir
                )
                patched_map[path] = tmp_file
            return patched_map

        workers = min(self.max_workers, len(dirty_edits))
        with ProcessPoolExecutor(max_workers=workers) as pool:
            futures = {
                pool.submit(
                    _worker_entrypoint,
                    patch_fn,
                    src_str,
                    sheet_path,
                    edits,
                    temp_dir,
                ): sheet_path
                for sheet_path, edits in dirty_edits.items()
            }

            for future in as_completed(futures):
                sheet_path, tmp_file = future.result()
                patched_map[sheet_path] = tmp_file

        return patched_map

    def _assemble_zip(
        self, patched_files: Mapping[str, str], strip_calc_chain: bool
    ) -> None:
        """Streams entries into the destination archive without buffering whole files."""
        dst_dir = self.dst_path.parent
        dst_dir.mkdir(parents=True, exist_ok=True)
        fd, staging_zip = tempfile.mkstemp(
            prefix=".tmp_xlsx_", suffix=".xlsx", dir=dst_dir
        )
        os.close(fd)

        try:
            with (
                zipfile.ZipFile(self.src_path, "r") as zin,
                zipfile.ZipFile(
                    staging_zip, "w", compression=zipfile.ZIP_DEFLATED
                ) as zout,
            ):
                for info in zin.infolist():
                    name = info.filename

                    if strip_calc_chain and name == CALC_CHAIN_PATH:
                        continue

                    if name in patched_files:
                        self._copy_disk_file_to_zip(zout, info, patched_files[name])
                        continue

                    if strip_calc_chain and name == "[Content_Types].xml":
                        cleaned = self._strip_calc_chain_content_type(zin.read(name))
                        zout.writestr(self._clone_zip_info(info), cleaned)
                        continue

                    if strip_calc_chain and name == "xl/_rels/workbook.xml.rels":
                        cleaned = self._strip_calc_chain_rel(zin.read(name))
                        zout.writestr(self._clone_zip_info(info), cleaned)
                        continue

                    if strip_calc_chain and name == "xl/workbook.xml":
                        updated = self._enable_full_calc_on_load(zin.read(name))
                        zout.writestr(self._clone_zip_info(info), updated)
                        continue

                    self._stream_zip_entry(zin, zout, info)

            os.replace(staging_zip, self.dst_path)
        except Exception:
            if os.path.exists(staging_zip):
                os.remove(staging_zip)
            raise

    def _stream_zip_entry(
        self, zin: zipfile.ZipFile, zout: zipfile.ZipFile, src_info: zipfile.ZipInfo
    ) -> None:
        """Copies a zip entry chunk-by-chunk with constant memory usage."""
        dst_info = self._clone_zip_info(src_info)
        with zin.open(src_info, "r") as src_f, zout.open(dst_info, "w") as dst_f:
            shutil.copyfileobj(src_f, dst_f, length=self.chunk_size)

    def _copy_disk_file_to_zip(
        self, zout: zipfile.ZipFile, src_info: zipfile.ZipInfo, disk_path: str
    ) -> None:
        """Streams a patched temporary XML file into the zip archive."""
        dst_info = self._clone_zip_info(src_info)
        with open(disk_path, "rb") as src_f, zout.open(dst_info, "w") as dst_f:
            shutil.copyfileobj(src_f, dst_f, length=self.chunk_size)

    @staticmethod
    def _clone_zip_info(src_info: zipfile.ZipInfo) -> zipfile.ZipInfo:
        """Preserves filename, timestamp, and compression type while resetting sizes."""
        info = zipfile.ZipInfo(
            filename=src_info.filename, date_time=src_info.date_time
        )
        info.compress_type = src_info.compress_type
        info.create_system = src_info.create_system
        info.external_attr = src_info.external_attr
        return info

    @staticmethod
    def _strip_calc_chain_content_type(xml_bytes: bytes) -> bytes:
        root = etree.fromstring(xml_bytes)
        for override in root.findall(f"{{{NS_CT}}}Override"):
            if (
                override.get("PartName") == "/xl/calcChain.xml"
                or override.get("ContentType") == CALC_CHAIN_CT
            ):
                root.remove(override)
        return etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )

    @staticmethod
    def _strip_calc_chain_rel(xml_bytes: bytes) -> bytes:
        root = etree.fromstring(xml_bytes)
        for rel in root.findall(f"{{{NS_REL_PKG}}}Relationship"):
            if rel.get("Type") == CALC_CHAIN_REL_TYPE:
                root.remove(rel)
        return etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )

    @staticmethod
    def _enable_full_calc_on_load(xml_bytes: bytes) -> bytes:
        root = etree.fromstring(xml_bytes)
        calc_pr = root.find(f"{{{NS_MAIN}}}calcPr")
        if calc_pr is None:
            calc_pr = etree.SubElement(root, f"{{{NS_MAIN}}}calcPr")
        calc_pr.set("fullCalcOnLoad", "1")
        return etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
