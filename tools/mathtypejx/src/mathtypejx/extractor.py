"""Extractor: extract MTEF byte data from MathType OLE binary objects."""

import io
import struct
from typing import Optional

import olefile

from .models import FormulaInfo, FormulaStatus
from .package import DEFAULT_LIMITS, PackageLimits

# Known stream names in MathType OLE containers
MTEF_STREAM_NAMES = [
    "Equation Native",    # Standard name (with space)
    "EquationNative",     # Alternative (no space)
    "Equation",           # Simplified
]

# EQNOLEFILEHDR structure: 28 bytes
#   WORD  cbHdr;       // size of header = 28
#   DWORD version;     // hiword=2, loword=0
#   WORD  cf;          // clipboard format
#   DWORD cbObject;    // length of MTEF data following header
#   DWORD reserved[4];
HEADER_STRUCT = struct.Struct("<H I H I I I I I")
HEADER_SIZE = HEADER_STRUCT.size  # 28 bytes


def extract_mtef(
    formula: FormulaInfo,
    ole_data: Optional[bytes] = None,
    limits: PackageLimits = DEFAULT_LIMITS,
) -> bool:
    """Extract MTEF byte data from an OLE binary and populate the FormulaInfo.

    Args:
        formula: FormulaInfo object. Must have ole_data set (or pass it directly).
        ole_data: Raw OLE binary bytes. If None, uses formula.ole_data.

    Returns:
        True if extraction succeeded, False otherwise.
    """
    data = ole_data or formula.ole_data
    if not data:
        formula.status = FormulaStatus.FAILED
        formula.error_message = "No OLE data available"
        return False
    if len(data) > limits.ole_part_max:
        formula.status = FormulaStatus.FAILED
        formula.error_message = "OLE data exceeds configured size limit"
        return False

    ole = None
    try:
        ole = olefile.OleFileIO(io.BytesIO(data))
    except Exception as e:
        formula.status = FormulaStatus.FAILED
        formula.error_message = f"Failed to open OLE container: {e}"
        return False

    mtef_bytes = None
    stream_name = None

    available_streams = [list(item) for item in ole.listdir()]
    try:
        for candidate in MTEF_STREAM_NAMES:
            if ole.exists(candidate):
                try:
                    stream_size = ole.get_size(candidate)
                    if stream_size > HEADER_SIZE + limits.mtef_stream_max:
                        continue
                    raw = ole.openstream(candidate).read(HEADER_SIZE + limits.mtef_stream_max + 1)
                    if len(raw) > HEADER_SIZE + limits.mtef_stream_max:
                        continue
                    mtef_bytes = _parse_header(raw, limits.mtef_stream_max)
                    stream_name = candidate
                    if mtef_bytes is not None:
                        break
                except (OSError, IOError, ValueError, struct.error):
                    continue
    finally:
        ole.close()

    if mtef_bytes is None:
        formula.status = FormulaStatus.FAILED
        formula.error_message = (
            f"No bounded MathType equation stream found. Available streams: {available_streams}"
        )
        return False

    formula.mtef_bytes = mtef_bytes
    formula.status = FormulaStatus.EXTRACTED
    return True


def _parse_header(raw: bytes, maximum_mtef_size: int = DEFAULT_LIMITS.mtef_stream_max) -> Optional[bytes]:
    """Parse the 28-byte EQNOLEFILEHDR and return the MTEF data portion.

    Returns None if the header is invalid.
    """
    if len(raw) < HEADER_SIZE:
        return None

    values = HEADER_STRUCT.unpack_from(raw, 0)
    cb_hdr, version, cf, cb_object, *_ = values

    if cb_hdr != HEADER_SIZE:
        return None

    mtef_data = raw[HEADER_SIZE:]
    if cb_object > 0:
        if cb_object > maximum_mtef_size or cb_object > len(mtef_data):
            return None
        mtef_data = mtef_data[:cb_object]
    elif len(mtef_data) > maximum_mtef_size:
        return None

    return mtef_data


def detect_mtef_version(mtef_bytes: bytes) -> Optional[int]:
    """Detect the MTEF version from the first byte.

    MTEF v5 header:
        Byte 0: MTEF version (5 for MathType 4.0+)
        Byte 1: Platform

    Returns the version number or None.
    """
    if not mtef_bytes or len(mtef_bytes) < 2:
        return None
    version = mtef_bytes[0]
    if version in (0, 1, 2, 3, 4, 5):
        return version
    return None


def inspect_ole(ole_data: bytes) -> dict:
    """Inspect an OLE container and return stream information (for debugging)."""
    info = {"streams": [], "header": None}
    try:
        ole = olefile.OleFileIO(io.BytesIO(ole_data))
        info["streams"] = [list(s) for s in ole.listdir()]

        for candidate in MTEF_STREAM_NAMES:
            if ole.exists(candidate):
                stream = ole.openstream(candidate)
                raw = stream.read()
                if len(raw) >= HEADER_SIZE:
                    values = HEADER_STRUCT.unpack_from(raw, 0)
                    info["header"] = {
                        "cbHdr": values[0],
                        "version": values[1],
                        "cf": values[2],
                        "cbObject": values[3],
                    }
                    mtef = raw[HEADER_SIZE:]
                    info["mtef_size"] = len(mtef)
                    info["mtef_version"] = detect_mtef_version(mtef)
                break
        ole.close()
    except Exception as e:
        info["error"] = str(e)
    return info
