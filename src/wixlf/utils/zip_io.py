import posixpath
import shutil
import time
import zipfile


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
) -> None:
    """
    Copies an untouched zip entry by transferring its raw compressed stream
    directly (avoiding CPU-heavy decompression and re-deflation).
    Falls back to standard decompressed copy for encrypted entries.
    """
    if src_info.flag_bits & 0x1:
        dst_info = clone_zip_info(src_info)
        with zin.open(src_info, "r") as src_f, zout.open(dst_info, "w") as dst_f:
            shutil.copyfileobj(src_f, dst_f, length=chunk_size)
        return

    raw_info = clone_zip_info(src_info)
    raw_info.CRC = src_info.CRC
    raw_info.compress_size = src_info.compress_size
    raw_info.file_size = src_info.file_size
    raw_info.flag_bits = src_info.flag_bits & ~0x8  # Clear data-descriptor bit

    with (
        zin.open(src_info, "r", raw=True) as src_f,
        zout.open(raw_info, "w", raw=True) as dst_f,
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
    with (
        open(disk_path, "rb") as src_f,
        zout.open(dst_info, "w", compresslevel=compresslevel) as dst_f,
    ):
        shutil.copyfileobj(src_f, dst_f, length=chunk_size)
