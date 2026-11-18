import posixpath
import shutil
import time
import zipfile
from pathlib import Path
from typing import NamedTuple

from ..constants import NS_REL_PKG
from .xml_text import parse_xml


class Relationship(NamedTuple):
    id: str
    type: str
    target: str
    external: bool


def read_relationships(zin: zipfile.ZipFile, rels_path: str) -> list[Relationship]:
    """Parses a .rels part into Relationship tuples (entries without Id/Target are skipped)."""
    rels_root = parse_xml(zin.read(rels_path))
    return [
        Relationship(
            id=rel_id,
            type=rel.get("Type", ""),
            target=target,
            external=rel.get("TargetMode") == "External",
        )
        for rel in rels_root.findall(f"{{{NS_REL_PKG}}}Relationship")
        if (rel_id := rel.get("Id")) and (target := rel.get("Target"))
    ]


def resolve_zip_path(base_file: str, target: str) -> str:
    """Resolves relative or root-anchored OOXML targets to zip entry names."""
    if target.startswith("/"):
        return target.lstrip("/")
    base_dir = posixpath.dirname(base_file)
    return posixpath.normpath(posixpath.join(base_dir, target))


def clone_zip_info(src_info: zipfile.ZipInfo) -> zipfile.ZipInfo:
    """Preserves filename, timestamp, and compression type while resetting sizes."""
    info = zipfile.ZipInfo(
        filename=src_info.filename, date_time=src_info.date_time
    )
    info.compress_type = src_info.compress_type
    info.create_system = src_info.create_system
    info.external_attr = src_info.external_attr
    return info


def new_zip_info(filename: str) -> zipfile.ZipInfo:
    """Creates a ZipInfo for a newly generated archive member with current timestamp."""
    now = time.localtime()[:6]
    info = zipfile.ZipInfo(filename=filename, date_time=now)
    info.compress_type = zipfile.ZIP_DEFLATED
    return info


def stream_zip_entry(
    zin: zipfile.ZipFile,
    zout: zipfile.ZipFile,
    src_info: zipfile.ZipInfo,
    chunk_size: int,
    compresslevel: int | None = 1,
) -> None:
    """
    Streams an untouched zip entry from zin to zout in constant O(1) memory
    using fast deflate compression (compresslevel=1 by default).
    """
    dst_info = clone_zip_info(src_info)
    if compresslevel is not None and dst_info.compress_type == zipfile.ZIP_DEFLATED:
        dst_info.compress_level = compresslevel
    zip64 = src_info.file_size > zipfile.ZIP64_LIMIT
    with (
        zin.open(src_info, "r") as src_f,
        zout.open(dst_info, "w", force_zip64=zip64) as dst_f,
    ):
        shutil.copyfileobj(src_f, dst_f, length=chunk_size)


def copy_disk_file_to_zip(
    zout: zipfile.ZipFile,
    target_info: zipfile.ZipInfo,
    disk_path: str,
    chunk_size: int,
    compresslevel: int | None = None,
) -> None:
    """Streams a temporary XML file from disk into the zip archive."""
    dst_info = clone_zip_info(target_info)
    if compresslevel is not None and dst_info.compress_type == zipfile.ZIP_DEFLATED:
        dst_info.compress_level = compresslevel
    zip64 = Path(disk_path).stat().st_size > zipfile.ZIP64_LIMIT
    with (
        open(disk_path, "rb") as src_f,
        zout.open(dst_info, "w", force_zip64=zip64) as dst_f,
    ):
        shutil.copyfileobj(src_f, dst_f, length=chunk_size)


def zip_entry_contains_formulas(
    zin: zipfile.ZipFile,
    zip_path: str,
    chunk_size: int = 1024 * 1024,
) -> bool:
    """
    Scans a worksheet XML stream in fixed-size chunks for '<f' or '<prefix:f'
    formula tags without building an XML DOM.
    """
    overlap = b""
    with zin.open(zip_path, "r") as f:
        while True:
            chunk = f.read(chunk_size)
            if not chunk:
                return False
            buf = overlap + chunk
            if (
                b"<f>" in buf
                or b"<f " in buf
                or b":f>" in buf
                or b":f " in buf
            ):
                return True
            overlap = buf[-8:]