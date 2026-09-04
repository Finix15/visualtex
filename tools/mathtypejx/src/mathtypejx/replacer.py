"""Transactional replacement of legacy OLE equations with validated OMML."""

from __future__ import annotations

import copy
import os
import posixpath
import shutil
import tempfile
import uuid
import zipfile
from pathlib import Path, PurePosixPath

from lxml import etree

from .models import FormulaInfo, FormulaStatus
from .package import (
    DEFAULT_LIMITS,
    PackageLimits,
    PackageValidationError,
    fsync_file_and_parent,
    parse_xml_file,
    resolve_relationship_target,
    safe_extract_docx,
    safe_package_path,
    safe_xml_parser,
)

NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_O = "urn:schemas-microsoft-com:office:office"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def replace_formulas(
    workdir: Path,
    formulas: list[FormulaInfo],
    output_docx: str,
    limits: PackageLimits = DEFAULT_LIMITS,
) -> list[FormulaInfo]:
    """Build and atomically publish a new DOCX; never mutate ``workdir``."""
    source_tree = Path(workdir).resolve()
    output_path = Path(output_docx).resolve()
    if output_path.suffix.lower() != ".docx":
        raise PackageValidationError("Output must use the .docx extension")
    if output_path.exists():
        raise FileExistsError(f"Output already exists: {output_path}")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    stage = Path(tempfile.mkdtemp(prefix=".visualtex-mathtype-stage-", dir=output_path.parent))
    candidate = output_path.parent / f".{output_path.name}.{uuid.uuid4()}.candidate.docx"
    replaced: set[str] = set()
    linked_output = False
    publish_complete = False
    try:
        if any(path.is_symlink() for path in source_tree.rglob("*")):
            raise PackageValidationError("Symlink in source package tree")
        shutil.copytree(source_tree, stage, dirs_exist_ok=True, symlinks=False)
        by_part: dict[str, list[FormulaInfo]] = {}
        for formula in formulas:
            if formula.status == FormulaStatus.CONVERTED:
                by_part.setdefault(formula.part_name, []).append(formula)

        for part_name, part_formulas in by_part.items():
            part_path = safe_package_path(stage, part_name)
            if not part_path.is_file():
                for formula in part_formulas:
                    formula.status = FormulaStatus.FAILED
                    formula.error_message = f"Part file not found: {part_name}"
                continue
            replaced.update(_replace_in_part(part_path, part_formulas, limits))

        _cleanup_relationships_and_embeddings(stage, formulas, replaced, limits)
        _repack_docx(stage, candidate)
        _validate_candidate(candidate, limits)
        fsync_file_and_parent(candidate)
        try:
            os.link(candidate, output_path)
            linked_output = True
        except FileExistsError:
            raise FileExistsError(f"Output already exists: {output_path}")
        fsync_file_and_parent(output_path)
        publish_complete = True
        for formula in formulas:
            if formula.formula_id in replaced:
                formula.status = FormulaStatus.REPLACED
        return formulas
    finally:
        if linked_output and not publish_complete and output_path.exists():
            output_path.unlink()
        if candidate.exists():
            candidate.unlink()
        shutil.rmtree(stage, ignore_errors=True)


def _parse_omml(formula: FormulaInfo):
    try:
        omml_xml = (formula.omml or "").strip()
        if omml_xml.startswith("<?xml"):
            end = omml_xml.find("?>")
            if end >= 0:
                omml_xml = omml_xml[end + 2:].strip()
        node = etree.fromstring(omml_xml.encode("utf-8"), parser=safe_xml_parser())
        qname = etree.QName(node)
        if qname.namespace != NS_M or qname.localname not in {"oMath", "oMathPara"}:
            raise ValueError(f"Unexpected OMML root: {node.tag}")
        return node
    except Exception as exc:
        formula.status = FormulaStatus.FAILED
        formula.error_message = f"OMML parse error: {exc}"
        return None


def _replace_in_part(
    part_path: Path,
    formulas: list[FormulaInfo],
    limits: PackageLimits = DEFAULT_LIMITS,
) -> set[str]:
    """Replace matching OLE children while preserving other content in each run."""
    tree = parse_xml_file(part_path, limits.xml_part_max)
    root = tree.getroot()
    rid_to_formula: dict[str, FormulaInfo] = {}
    rid_to_omml = {}
    for formula in formulas:
        if formula.status != FormulaStatus.CONVERTED:
            continue
        if formula.relationship_id in rid_to_formula:
            formula.status = FormulaStatus.FAILED
            formula.error_message = "Duplicate converted relationship id in one part"
            continue
        node = _parse_omml(formula)
        if node is not None:
            rid_to_formula[formula.relationship_id] = formula
            rid_to_omml[formula.relationship_id] = node

    replaced: set[str] = set()
    for run in list(root.iter(f"{{{NS_W}}}r")):
        parent = run.getparent()
        if parent is None:
            continue
        run_properties = run.find(f"{{{NS_W}}}rPr")
        segments: list[object] = []
        buffered_children = []
        run_replaced = False

        def flush_buffer() -> None:
            if not buffered_children:
                return
            clone = etree.Element(run.tag, nsmap=run.nsmap)
            for key, value in run.attrib.items():
                clone.set(key, value)
            if run_properties is not None:
                clone.append(copy.deepcopy(run_properties))
            for child in buffered_children:
                clone.append(copy.deepcopy(child))
            segments.append(clone)
            buffered_children.clear()

        for child in run:
            if child is run_properties:
                continue
            matching = []
            for ole in child.iter(f"{{{NS_O}}}OLEObject"):
                rid = ole.get(f"{{{NS_R}}}id", "")
                if rid in rid_to_formula:
                    matching.append(rid)
            if len(matching) == 1:
                flush_buffer()
                rid = matching[0]
                segments.append(copy.deepcopy(rid_to_omml[rid]))
                replaced.add(rid_to_formula[rid].formula_id)
                run_replaced = True
            elif len(matching) > 1:
                for rid in matching:
                    formula = rid_to_formula[rid]
                    formula.status = FormulaStatus.FAILED
                    formula.error_message = "Multiple converted OLE objects share one run child"
                buffered_children.append(child)
            else:
                buffered_children.append(child)
        flush_buffer()
        if not run_replaced:
            continue
        original_index = parent.index(run)
        parent.remove(run)
        for offset, segment in enumerate(segments):
            parent.insert(original_index + offset, segment)

    if replaced:
        part_path.write_bytes(etree.tostring(
            root,
            encoding="utf-8",
            xml_declaration=True,
            standalone=True,
            pretty_print=False,
        ))
    return replaced


def _source_part_for_rels(rels_name: str) -> str:
    if rels_name == "_rels/.rels":
        return ""
    path = PurePosixPath(rels_name)
    if path.parent.name != "_rels" or not path.name.endswith(".rels"):
        raise PackageValidationError(f"Invalid relationships part name: {rels_name}")
    return posixpath.join(str(path.parent.parent), path.name[:-5])


def _relationship_targets(workdir: Path, limits: PackageLimits) -> set[str]:
    targets: set[str] = set()
    for rels_path in workdir.rglob("*.rels"):
        rels_name = rels_path.relative_to(workdir).as_posix()
        source_part = _source_part_for_rels(rels_name)
        tree = parse_xml_file(rels_path, limits.xml_part_max)
        for rel in tree.getroot():
            if rel.tag != f"{{{PACKAGE_REL_NS}}}Relationship":
                continue
            if rel.get("TargetMode", "").lower() == "external":
                continue
            try:
                targets.add(resolve_relationship_target(source_part, rel.get("Target", "")))
            except PackageValidationError:
                continue
    return targets


def _cleanup_relationships_and_embeddings(
    workdir: Path,
    formulas: list[FormulaInfo],
    replaced: set[str],
    limits: PackageLimits,
) -> None:
    by_part: dict[str, list[FormulaInfo]] = {}
    for formula in formulas:
        if formula.formula_id in replaced:
            by_part.setdefault(formula.part_name, []).append(formula)
    candidate_embeddings: set[str] = set()
    for part_name, part_formulas in by_part.items():
        part_tree = parse_xml_file(safe_package_path(workdir, part_name), limits.xml_part_max)
        remaining_ids = {
            value for element in part_tree.getroot().iter()
            for name, value in element.attrib.items()
            if etree.QName(name).namespace == NS_R and etree.QName(name).localname == "id"
        }
        rels_path = safe_package_path(workdir, part_formulas[0].rels_path)
        if not rels_path.exists():
            continue
        rels_tree = parse_xml_file(rels_path, limits.xml_part_max)
        changed = False
        removable_ids = {f.relationship_id for f in part_formulas} - remaining_ids
        for rel in list(rels_tree.getroot()):
            if rel.get("Id", "") not in removable_ids:
                continue
            try:
                candidate_embeddings.add(resolve_relationship_target(part_name, rel.get("Target", "")))
            except PackageValidationError:
                pass
            rels_tree.getroot().remove(rel)
            changed = True
        if changed:
            rels_path.write_bytes(etree.tostring(
                rels_tree.getroot(), encoding="utf-8", xml_declaration=True, standalone=True
            ))

    referenced_targets = _relationship_targets(workdir, limits)
    deleted_parts: set[str] = set()
    for package_name in candidate_embeddings - referenced_targets:
        target = safe_package_path(workdir, package_name)
        if target.is_file() and not target.is_symlink():
            target.unlink()
            deleted_parts.add(package_name)
    _remove_deleted_content_type_overrides(workdir, deleted_parts, limits)


def _remove_deleted_content_type_overrides(
    workdir: Path,
    deleted_parts: set[str],
    limits: PackageLimits,
) -> None:
    if not deleted_parts:
        return
    content_types = safe_package_path(workdir, "[Content_Types].xml")
    tree = parse_xml_file(content_types, limits.xml_part_max)
    root = tree.getroot()
    namespace = "http://schemas.openxmlformats.org/package/2006/content-types"
    deleted_names = {f"/{name}" for name in deleted_parts}
    changed = False
    for override in list(root):
        if override.tag == f"{{{namespace}}}Override" and override.get("PartName") in deleted_names:
            root.remove(override)
            changed = True
    if changed:
        content_types.write_bytes(etree.tostring(
            root, encoding="utf-8", xml_declaration=True, standalone=True
        ))


def _validate_candidate(candidate: Path, limits: PackageLimits) -> None:
    verify_root = Path(tempfile.mkdtemp(prefix="visualtex-mathtype-verify-"))
    try:
        safe_extract_docx(candidate, verify_root, limits)
        document = safe_package_path(verify_root, "word/document.xml")
        parse_xml_file(document, limits.xml_part_max)
    finally:
        shutil.rmtree(verify_root, ignore_errors=True)


def _update_content_types(workdir: Path) -> None:
    """Compatibility no-op: OMML is inline XML and needs no content-type override."""
    del workdir


def _repack_docx(workdir: Path, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output_path, "x", zipfile.ZIP_DEFLATED, allowZip64=True) as archive:
        for file_path in sorted(workdir.rglob("*")):
            if file_path.is_symlink():
                raise PackageValidationError(f"Symlink in staging tree: {file_path}")
            if file_path.is_file():
                archive.write(file_path, file_path.relative_to(workdir).as_posix())
