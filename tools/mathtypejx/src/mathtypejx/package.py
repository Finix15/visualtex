"""Security and OPC helpers for Word Open XML packages."""

from __future__ import annotations

import os
import posixpath
import stat
import zipfile
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from urllib.parse import unquote, urlsplit

from lxml import etree


class PackageValidationError(ValueError):
    """Raised when an input is not a safe, supported DOCX package."""


@dataclass(frozen=True)
class PackageLimits:
    compressed_docx_max: int = 500 * 1024 * 1024
    uncompressed_package_max: int = 2 * 1024 * 1024 * 1024
    zip_entries_max: int = 50_000
    xml_part_max: int = 64 * 1024 * 1024
    ole_part_max: int = 32 * 1024 * 1024
    formula_count_max: int = 20_000
    mtef_stream_max: int = 8 * 1024 * 1024


DEFAULT_LIMITS = PackageLimits()


def safe_xml_parser() -> etree.XMLParser:
    return etree.XMLParser(
        resolve_entities=False,
        load_dtd=False,
        no_network=True,
        recover=False,
        huge_tree=False,
        remove_blank_text=False,
    )


def parse_xml_file(path: Path, maximum_size: int) -> etree._ElementTree:
    if not path.is_file() or path.is_symlink():
        raise PackageValidationError(f"XML part is not a regular file: {path}")
    size = path.stat().st_size
    if size > maximum_size:
        raise PackageValidationError(f"XML part exceeds limit ({size} > {maximum_size}): {path}")
    data = path.read_bytes()
    if b"<!DOCTYPE" in data.upper():
        raise PackageValidationError(f"DTD is not allowed in XML part: {path}")
    try:
        return etree.ElementTree(etree.fromstring(data, parser=safe_xml_parser()))
    except (etree.XMLSyntaxError, ValueError) as exc:
        raise PackageValidationError(f"Malformed XML part {path}: {exc}") from exc


def normalize_package_name(name: str) -> str:
    """Validate a ZIP member name and return its canonical OPC name."""
    decoded = unquote(name)
    if not decoded or "\x00" in decoded or "\\" in decoded:
        raise PackageValidationError(f"Unsafe ZIP entry name: {name!r}")
    parsed = urlsplit(decoded)
    if parsed.scheme or parsed.netloc or parsed.query or parsed.fragment:
        raise PackageValidationError(f"Unsafe ZIP entry URI: {name!r}")
    if decoded.startswith("/"):
        raise PackageValidationError(f"Absolute ZIP entry is not allowed: {name!r}")
    parts = PurePosixPath(decoded).parts
    if any(part in ("", ".", "..") for part in parts):
        raise PackageValidationError(f"ZIP traversal entry is not allowed: {name!r}")
    normalized = posixpath.normpath(decoded)
    if normalized == ".." or normalized.startswith("../"):
        raise PackageValidationError(f"ZIP traversal entry is not allowed: {name!r}")
    return normalized


def resolve_relationship_target(source_part: str, target: str) -> str:
    """Resolve an internal OPC relationship target without filesystem semantics."""
    decoded = unquote(target)
    if not decoded or "\x00" in decoded or "\\" in decoded:
        raise PackageValidationError(f"Unsafe relationship target: {target!r}")
    parsed = urlsplit(decoded)
    if parsed.scheme or parsed.netloc or parsed.query or parsed.fragment:
        raise PackageValidationError(f"External relationship target is not allowed: {target!r}")
    if decoded.startswith("/"):
        candidate = decoded.lstrip("/")
    else:
        candidate = posixpath.join(posixpath.dirname(source_part), decoded)
    normalized = posixpath.normpath(candidate)
    if normalized in ("", ".", "..") or normalized.startswith("../"):
        raise PackageValidationError(f"Relationship escapes package root: {target!r}")
    return normalize_package_name(normalized)


def safe_package_path(root: Path, package_name: str) -> Path:
    canonical = normalize_package_name(package_name)
    root_resolved = root.resolve()
    candidate = root_resolved.joinpath(*PurePosixPath(canonical).parts)
    parent = candidate.parent.resolve(strict=False)
    if parent != root_resolved and root_resolved not in parent.parents:
        raise PackageValidationError(f"Package path escapes extraction root: {package_name!r}")
    return candidate


def _is_zip_symlink(info: zipfile.ZipInfo) -> bool:
    return stat.S_IFMT(info.external_attr >> 16) == stat.S_IFLNK


def safe_extract_docx(
    docx_path: Path,
    workdir: Path,
    limits: PackageLimits = DEFAULT_LIMITS,
) -> None:
    """Validate and extract a DOCX without trusting ZIP member paths or sizes."""
    if docx_path.suffix.lower() != ".docx":
        raise PackageValidationError("Only .docx input is supported")
    if docx_path.is_symlink() or not docx_path.is_file():
        raise PackageValidationError("Input must be a regular file, not a symlink")
    compressed_size = docx_path.stat().st_size
    if compressed_size > limits.compressed_docx_max:
        raise PackageValidationError("Compressed DOCX exceeds configured limit")
    if not zipfile.is_zipfile(docx_path):
        raise PackageValidationError("Input is not a valid ZIP/DOCX package")

    with zipfile.ZipFile(docx_path, "r") as archive:
        infos = archive.infolist()
        if len(infos) > limits.zip_entries_max:
            raise PackageValidationError("DOCX contains too many ZIP entries")
        seen: set[str] = set()
        total_declared = 0
        validated: list[tuple[zipfile.ZipInfo, str]] = []
        for info in infos:
            canonical = normalize_package_name(info.filename.rstrip("/"))
            if canonical in seen:
                raise PackageValidationError(f"Duplicate ZIP entry: {canonical}")
            seen.add(canonical)
            if info.flag_bits & 0x1:
                raise PackageValidationError(f"Encrypted ZIP entry is not supported: {canonical}")
            if _is_zip_symlink(info):
                raise PackageValidationError(f"ZIP symlink is not allowed: {canonical}")
            total_declared += info.file_size
            if total_declared > limits.uncompressed_package_max:
                raise PackageValidationError("Uncompressed DOCX exceeds configured limit")
            lower = canonical.lower()
            if lower.startswith("_xmlsignatures/") or lower == "_xmlsignatures":
                raise PackageValidationError("Digitally signed DOCX is blocked in the MVP")
            if lower.endswith((".xml", ".rels")) and info.file_size > limits.xml_part_max:
                raise PackageValidationError(f"XML part exceeds configured limit: {canonical}")
            if "/embeddings/" in f"/{lower}" and info.file_size > limits.ole_part_max:
                raise PackageValidationError(f"OLE part exceeds configured limit: {canonical}")
            validated.append((info, canonical))

        for info, canonical in validated:
            destination = safe_package_path(workdir, canonical)
            if info.is_dir():
                destination.mkdir(parents=True, exist_ok=True)
                continue
            destination.parent.mkdir(parents=True, exist_ok=True)
            written = 0
            with archive.open(info, "r") as source, destination.open("xb") as target_file:
                while True:
                    chunk = source.read(1024 * 1024)
                    if not chunk:
                        break
                    written += len(chunk)
                    if written > info.file_size:
                        raise PackageValidationError(f"ZIP entry exceeded declared size: {canonical}")
                    target_file.write(chunk)
            if written != info.file_size:
                raise PackageValidationError(f"ZIP entry size mismatch: {canonical}")

    for required in ("[Content_Types].xml", "_rels/.rels"):
        if not safe_package_path(workdir, required).is_file():
            raise PackageValidationError(f"Required OPC part is missing: {required}")


def fsync_file_and_parent(path: Path) -> None:
    with path.open("rb") as handle:
        os.fsync(handle.fileno())
    directory_fd = os.open(path.parent, os.O_RDONLY)
    try:
        os.fsync(directory_fd)
    finally:
        os.close(directory_fd)
