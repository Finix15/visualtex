"""Safe DOCX/OPC scanner for MathType and legacy Equation Editor OLE objects."""

from __future__ import annotations

import posixpath
import tempfile
from collections import deque
from pathlib import Path, PurePosixPath
from typing import Optional

from .models import FormulaInfo, FormulaType
from .package import (
    DEFAULT_LIMITS,
    PackageLimits,
    PackageValidationError,
    parse_xml_file,
    resolve_relationship_target,
    safe_extract_docx,
    safe_package_path,
)

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "o": "urn:schemas-microsoft-com:office:office",
}
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE_DOCUMENT_REL = "/officeDocument"
OLE_OBJECT_REL = "/oleObject"
STORY_REL_SUFFIXES = {"/header", "/footer", "/footnotes", "/endnotes"}
MATHTYPE_PROG_IDS = {"Equation.DSMT4", "Equation.3", "Equation.2"}


def _rels_path_for_part(part_name: str) -> str:
    name = PurePosixPath(part_name).name
    parent = posixpath.dirname(part_name)
    return posixpath.join(parent, "_rels", f"{name}.rels")


def _read_relationships(
    workdir: Path,
    source_part: str,
    limits: PackageLimits,
) -> list[dict[str, str]]:
    rels_name = "_rels/.rels" if not source_part else _rels_path_for_part(source_part)
    rels_path = safe_package_path(workdir, rels_name)
    if not rels_path.exists():
        return []
    tree = parse_xml_file(rels_path, limits.xml_part_max)
    relationships = []
    for rel in tree.getroot():
        if rel.tag != f"{{{PACKAGE_REL_NS}}}Relationship":
            continue
        relationships.append({
            "id": rel.get("Id", ""),
            "type": rel.get("Type", ""),
            "target": rel.get("Target", ""),
            "target_mode": rel.get("TargetMode", ""),
        })
    return relationships


def _discover_word_parts(workdir: Path, limits: PackageLimits) -> list[str]:
    root_rels = _read_relationships(workdir, "", limits)
    office_targets = [
        rel for rel in root_rels
        if rel["type"].endswith(OFFICE_DOCUMENT_REL)
        and rel["target_mode"].lower() != "external"
    ]
    if len(office_targets) != 1:
        raise PackageValidationError("DOCX must contain exactly one officeDocument relationship")
    main_part = resolve_relationship_target("", office_targets[0]["target"])
    queue = deque([main_part])
    discovered: list[str] = []
    seen: set[str] = set()
    while queue:
        part_name = queue.popleft()
        if part_name in seen:
            continue
        seen.add(part_name)
        part_path = safe_package_path(workdir, part_name)
        if not part_path.is_file():
            raise PackageValidationError(f"WordprocessingML part is missing: {part_name}")
        discovered.append(part_name)
        for rel in _read_relationships(workdir, part_name, limits):
            if rel["target_mode"].lower() == "external":
                continue
            if any(rel["type"].endswith(suffix) for suffix in STORY_REL_SUFFIXES):
                queue.append(resolve_relationship_target(part_name, rel["target"]))
    return discovered


def _relationships_by_id(
    workdir: Path,
    part_name: str,
    limits: PackageLimits,
) -> dict[str, dict[str, str]]:
    result: dict[str, dict[str, str]] = {}
    for rel in _read_relationships(workdir, part_name, limits):
        if rel["target_mode"].lower() == "external" or not rel["id"]:
            continue
        if rel["id"] in result:
            raise PackageValidationError(
                f"Duplicate relationship id {rel['id']!r} in {_rels_path_for_part(part_name)}"
            )
        target = resolve_relationship_target(part_name, rel["target"])
        result[rel["id"]] = {**rel, "resolved_target": target}
    return result


def scan_docx(
    docx_path: str,
    workdir: Optional[str] = None,
    limits: PackageLimits = DEFAULT_LIMITS,
) -> tuple[list[FormulaInfo], Path]:
    """Validate, safely extract and scan a DOCX using OPC relationships."""
    source = Path(docx_path).absolute()
    if workdir:
        wd = Path(workdir)
        if wd.exists() and any(wd.iterdir()):
            raise PackageValidationError("Extraction workdir must be empty")
        wd.mkdir(parents=True, exist_ok=True)
    else:
        wd = Path(tempfile.mkdtemp(prefix="mathtypejx_"))

    safe_extract_docx(source, wd, limits)
    formulas: list[FormulaInfo] = []
    for part_name in _discover_word_parts(wd, limits):
        part_path = safe_package_path(wd, part_name)
        tree = parse_xml_file(part_path, limits.xml_part_max)
        relationships = _relationships_by_id(wd, part_name, limits)
        for ole_elem in part_path_ole_objects(tree, part_name):
            if len(formulas) >= limits.formula_count_max:
                raise PackageValidationError("Formula count exceeds configured limit")
            rid = ole_elem.get(f"{{{NS['r']}}}id", "")
            prog_id = ole_elem.get("ProgID", "")
            relationship = relationships.get(rid)
            ole_part_name = ""
            if relationship and relationship["type"].endswith(OLE_OBJECT_REL):
                ole_part_name = relationship["resolved_target"]
            ole_name = (
                PurePosixPath(ole_part_name).name
                if ole_part_name
                else f"oleObject_unknown_{len(formulas)}.bin"
            )
            para_idx, run_idx = _find_location(ole_elem)
            formula_type = FormulaType.MATHTYPE_OLE
            if prog_id in {"Equation.3", "Equation.2"}:
                formula_type = FormulaType.EQUATION_EDITOR_3
            formulas.append(FormulaInfo(
                formula_id=f"F{len(formulas) + 1:04d}",
                ole_name=ole_name,
                part_name=part_name,
                rels_path=_rels_path_for_part(part_name),
                relationship_id=rid,
                prog_id=prog_id,
                para_index=para_idx,
                run_index=run_idx,
                formula_type=formula_type,
                ole_part_name=ole_part_name,
            ))
    return formulas, wd


def part_path_ole_objects(tree, part_name: str):
    del part_name
    for ole in tree.iter():
        if ole.tag == f"{{{NS['o']}}}OLEObject" and ole.get("ProgID", "") in MATHTYPE_PROG_IDS:
            yield ole


def _find_location(elem) -> tuple[int, int]:
    para_idx = -1
    run_idx = -1
    current = elem
    while current is not None:
        tag = current.tag
        if tag == f"{{{NS['w']}}}r" and run_idx == -1:
            parent = current.getparent()
            if parent is not None:
                run_idx = parent.index(current)
        elif tag == f"{{{NS['w']}}}p":
            parent = current.getparent()
            if parent is not None:
                para_idx = parent.index(current)
            break
        current = current.getparent()
    return para_idx, run_idx


def read_ole_binary(
    workdir: Path,
    ole_name: str,
    ole_part_name: Optional[str] = None,
    limits: PackageLimits = DEFAULT_LIMITS,
) -> Optional[bytes]:
    """Read a relationship-resolved OLE part without escaping the package root."""
    package_name = ole_part_name or posixpath.join("word", "embeddings", ole_name)
    try:
        ole_path = safe_package_path(Path(workdir), package_name)
    except PackageValidationError:
        return None
    if ole_path.is_symlink() or not ole_path.is_file():
        return None
    if ole_path.stat().st_size > limits.ole_part_max:
        return None
    return ole_path.read_bytes()
