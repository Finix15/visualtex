"""Bounded file-protocol worker used by the macOS VisualTeX application.

This entrypoint deliberately stops at MathML during scan.  MathML to OMML is
performed by the TypeScript converter and only validated OMML is accepted back
for transactional DOCX replacement.
"""

from __future__ import annotations

import argparse
import base64
import hashlib
import json
import os
import shutil
import sys
import tempfile
from pathlib import Path
from typing import Any

from lxml import etree

from mathtypejx.converter import _normalize_mathml
from mathtypejx.extractor import detect_mtef_version, extract_mtef
from mathtypejx.models import FormulaInfo, FormulaStatus, RiskLevel
from mathtypejx.mtef import mtef_to_mathml
from mathtypejx.package import safe_xml_parser
from mathtypejx.replacer import replace_formulas
from mathtypejx.scanner import read_ole_binary, scan_docx

PROTOCOL_VERSION = 1
WORKER_VERSION = "0.1.0"
MAX_MANIFEST_BYTES = 4 * 1024 * 1024
MAX_BATCH_BYTES = 16 * 1024 * 1024
MAX_BATCH_FORMULAS = 64
MAX_MATHML_BYTES = 512 * 1024
MAX_OMML_BYTES = 1024 * 1024
MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def _read_json(path: Path, maximum: int) -> dict[str, Any]:
    st = path.lstat()
    if path.is_symlink() or not path.is_file() or st.st_size > maximum:
        raise ValueError(f"Unsafe or oversized protocol file: {path.name}")
    value = json.loads(path.read_bytes())
    if not isinstance(value, dict):
        raise ValueError(f"Protocol file must contain an object: {path.name}")
    return value


def _write_json(path: Path, value: dict[str, Any], maximum: int) -> None:
    data = json.dumps(value, ensure_ascii=False, indent=2).encode("utf-8")
    if len(data) > maximum:
        raise ValueError(f"Protocol output exceeds limit: {path.name}")
    temporary = path.with_name(f".{path.name}.tmp")
    flags = os.O_WRONLY | os.O_CREAT | os.O_EXCL
    descriptor = os.open(temporary, flags, 0o600)
    try:
        with os.fdopen(descriptor, "wb") as stream:
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
        directory = os.open(path.parent, os.O_RDONLY)
        try:
            os.fsync(directory)
        finally:
            os.close(directory)
    finally:
        if temporary.exists():
            temporary.unlink()


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(64 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _risk(mathml: str) -> RiskLevel:
    try:
        root = etree.fromstring(mathml.encode("utf-8"))
    except (ValueError, etree.XMLSyntaxError):
        return RiskLevel.BLOCKED
    names = {etree.QName(node).localname for node in root.iter() if isinstance(node.tag, str)}
    if names & {"mmultiscripts", "mprescripts", "mtable", "mtr", "mtd", "mroot", "menclose"}:
        return RiskLevel.MANUAL_REVIEW
    if names & {"mfrac", "msqrt", "msub", "msup", "msubsup", "munder", "mover", "munderover"}:
        return RiskLevel.SPOT_CHECK
    return RiskLevel.AUTO_REPLACE


def _public_prog_id(value: str) -> str:
    return value if value in {"Equation.DSMT4", "Equation.3", "Equation.2"} else "unknown"


def _formula_payload(formula: FormulaInfo, workdir: Path) -> dict[str, Any]:
    errors: list[str] = []
    warnings: list[str] = []
    formula.ole_data = read_ole_binary(workdir, formula.ole_name, formula.ole_part_name)
    if not formula.ole_data or not extract_mtef(formula):
        errors.append(formula.error_message or "MTEF extraction failed")
    else:
        formula.mtef_version = detect_mtef_version(formula.mtef_bytes or b"")
        try:
            formula.mathml = _normalize_mathml(mtef_to_mathml(formula.ole_data))
            if len(formula.mathml.encode("utf-8")) > MAX_MATHML_BYTES:
                raise ValueError("MathML exceeds configured size limit")
            formula.risk_level = _risk(formula.mathml)
        except Exception as error:  # parser errors are formula-local
            formula.mathml = None
            formula.risk_level = RiskLevel.BLOCKED
            formula.status = FormulaStatus.FAILED
            errors.append(f"MTEF to MathML conversion failed: {error}")
    if formula.mtef_version not in {3, 5}:
        warnings.append("MTEF version is not validated by the current corpus")
        if formula.mtef_version != 5:
            formula.risk_level = RiskLevel.BLOCKED
    risk = formula.risk_level.value.replace("_", "-")
    return {
        "formulaId": formula.formula_id,
        "partName": formula.part_name,
        "relationshipId": formula.relationship_id,
        "olePartName": formula.ole_part_name,
        "progId": _public_prog_id(formula.prog_id),
        "mtefVersion": formula.mtef_version if formula.mtef_version in {3, 5} else None,
        "displayMode": "inline" if formula.is_inline else "block",
        "status": "failed" if errors else "extracted",
        "riskLevel": "blocked" if errors else risk,
        **({"mathMl": formula.mathml} if formula.mathml else {}),
        "warnings": warnings,
        "errors": errors,
    }


def scan(job_id: str, job_root: Path) -> None:
    request = _read_json(job_root / "request.json", MAX_MANIFEST_BYTES)
    if request.get("protocolVersion") != PROTOCOL_VERSION or request.get("jobId") != job_id:
        raise ValueError("Request does not match worker protocol or job")
    source = Path(request.get("inputPath", ""))
    expected_hash = request.get("inputSha256", "")
    if not source.is_absolute() or source.is_symlink() or _sha256(source) != expected_hash:
        raise ValueError("Source document is unsafe or its hash changed")
    temporary = Path(tempfile.mkdtemp(prefix="visualtex-mathtype-scan-"))
    try:
        formulas, workdir = scan_docx(str(source), str(temporary / "package"))
        payloads = [_formula_payload(formula, workdir) for formula in formulas]
    finally:
        shutil.rmtree(temporary, ignore_errors=True)
    batch_count = (len(payloads) + MAX_BATCH_FORMULAS - 1) // MAX_BATCH_FORMULAS
    for index in range(batch_count):
        envelope = {
            "protocolVersion": PROTOCOL_VERSION,
            "jobId": job_id,
            "batchIndex": index,
            "batchCount": batch_count,
            "formulas": payloads[index * MAX_BATCH_FORMULAS:(index + 1) * MAX_BATCH_FORMULAS],
        }
        _write_json(job_root / f"formula-batch-{index}.json", envelope, MAX_BATCH_BYTES)
    by_prog: dict[str, int] = {}
    by_risk: dict[str, int] = {}
    for formula in payloads:
        by_prog[formula["progId"]] = by_prog.get(formula["progId"], 0) + 1
        by_risk[formula["riskLevel"]] = by_risk.get(formula["riskLevel"], 0) + 1
    _write_json(job_root / "scan-report.json", {
        "protocolVersion": PROTOCOL_VERSION, "jobId": job_id,
        "detected": len(payloads), "batchCount": batch_count,
        "batchSize": MAX_BATCH_FORMULAS, "byProgId": by_prog, "byRiskLevel": by_risk,
    }, MAX_MANIFEST_BYTES)


def _decode_omml(value: str) -> str:
    if not isinstance(value, str) or len(value) > (MAX_OMML_BYTES * 2):
        raise ValueError("OMML payload is missing or oversized")
    padding = "=" * (-len(value) % 4)
    raw = base64.urlsafe_b64decode(value + padding)
    if len(raw) > MAX_OMML_BYTES:
        raise ValueError("OMML payload is oversized")
    text = raw.decode("utf-8")
    node = etree.fromstring(raw, parser=safe_xml_parser())
    name = etree.QName(node)
    if name.namespace != MATH_NS or name.localname not in {"oMath", "oMathPara"}:
        raise ValueError("OMML root is invalid")
    return text


def finalize(job_id: str, job_root: Path) -> None:
    request = _read_json(job_root / "request.json", MAX_MANIFEST_BYTES)
    report = _read_json(job_root / "scan-report.json", MAX_MANIFEST_BYTES)
    source = Path(request.get("inputPath", ""))
    expected_hash = request.get("inputSha256", "")
    if request.get("jobId") != job_id or report.get("jobId") != job_id or _sha256(source) != expected_hash:
        raise ValueError("Finalize source or job mismatch")
    decisions: dict[str, dict[str, Any]] = {}
    batch_count = report.get("batchCount")
    if not isinstance(batch_count, int) or batch_count < 0:
        raise ValueError("Invalid scan batch count")
    for index in range(batch_count):
        batch = _read_json(job_root / f"omml-batch-{index}.json", MAX_BATCH_BYTES)
        if batch.get("jobId") != job_id or batch.get("batchIndex") != index or batch.get("batchCount") != batch_count:
            raise ValueError("OMML batch does not match job ordering")
        batch_formulas = batch.get("formulas")
        if not isinstance(batch_formulas, list):
            raise ValueError("OMML batch formulas must be an array")
        for result in batch_formulas:
            formula_id = result.get("formulaId") if isinstance(result, dict) else None
            if not isinstance(formula_id, str) or formula_id in decisions:
                raise ValueError("Missing or duplicate formula result")
            decisions[formula_id] = result
    temporary = Path(tempfile.mkdtemp(prefix="visualtex-mathtype-finalize-"))
    candidate = job_root / "candidate.docx"
    try:
        formulas, workdir = scan_docx(str(source), str(temporary / "package"))
        if set(decisions) != {formula.formula_id for formula in formulas}:
            raise ValueError("Formula count conservation failed")
        for formula in formulas:
            result = decisions[formula.formula_id]
            if result.get("status") == "replaced" and not result.get("errors"):
                try:
                    formula.omml = _decode_omml(result.get("ommlBase64"))
                    formula.status = FormulaStatus.CONVERTED
                except Exception as error:
                    formula.status = FormulaStatus.FAILED
                    formula.error_message = str(error)
            else:
                formula.status = FormulaStatus.SKIPPED
        updated = replace_formulas(workdir, formulas, str(candidate))
    finally:
        shutil.rmtree(temporary, ignore_errors=True)
    results = []
    counts = {"replaced": 0, "preserved": 0, "skipped": 0, "failed": 0}
    for formula in updated:
        requested = decisions[formula.formula_id]
        if formula.status == FormulaStatus.REPLACED:
            status = "replaced"
        elif requested.get("status") == "preserved":
            status = "preserved"
        elif formula.status == FormulaStatus.FAILED:
            status = "failed"
        else:
            status = "skipped"
        counts[status] += 1
        errors = list(requested.get("errors") or [])
        if formula.error_message:
            errors.append(formula.error_message)
        results.append({"formulaId": formula.formula_id, "status": status,
                        "warnings": list(requested.get("warnings") or []), "errors": errors})
    output_hash = _sha256(candidate)
    _write_json(job_root / "conversion-report.json", {
        "detected": len(updated), **counts, "sourceUnmodified": _sha256(source) == expected_hash,
        "inputSha256": expected_hash, "outputSha256": output_hash,
        "packageValid": True, "formulas": results,
    }, MAX_MANIFEST_BYTES)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--protocol", type=int, required=True)
    parser.add_argument("--job-id", required=True)
    parser.add_argument("--job-root", type=Path, required=True)
    parser.add_argument("--operation", choices=("scan", "finalize"), required=True)
    args = parser.parse_args(argv)
    try:
        if args.protocol != PROTOCOL_VERSION or args.job_root.is_symlink() or not args.job_root.is_dir():
            raise ValueError("Unsupported or unsafe worker invocation")
        (scan if args.operation == "scan" else finalize)(args.job_id, args.job_root.resolve())
        status = "awaitingOmml" if args.operation == "scan" else "complete"
        print(json.dumps({"protocolVersion": PROTOCOL_VERSION, "jobId": args.job_id, "status": status}))
        return 0
    except Exception as error:
        print(f"legacy-equation worker failed: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
