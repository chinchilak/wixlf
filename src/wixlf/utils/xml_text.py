from datetime import date, datetime, time
from typing import Final

from lxml import etree

_SAFE_PARSER: Final[etree.XMLParser] = etree.XMLParser(
    resolve_entities=False, no_network=True, load_dtd=False
)


def reject_doctype(root: etree._Element) -> None:
    """Raises if the document declares a DTD (never present in genuine OOXML parts)."""
    if root.getroottree().docinfo.doctype:
        raise ValueError("DTD declarations are not allowed in OOXML parts.")


def parse_xml(data: bytes) -> etree._Element:
    """Parses an OOXML part with entity expansion, network access, and DTDs disabled."""
    root = etree.fromstring(data, parser=_SAFE_PARSER)
    reject_doctype(root)
    return root

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


def to_excel_serial(dt_val: datetime | date | time) -> float:
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
    if not isinstance(dt_val, time) and serial < 61:
        serial -= 1.0
    return serial
