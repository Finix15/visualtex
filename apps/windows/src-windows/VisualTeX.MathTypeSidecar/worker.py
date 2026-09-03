"""VisualTeX MathType sidecar protocol v1 (offline, no Office automation)."""
from __future__ import annotations
import hashlib, json, os, sys
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from mathtypejx.converter import convert_formula
from mathtypejx.extractor import detect_mtef_version, extract_mtef
from mathtypejx.models import FormulaInfo, FormulaStatus

PROTOCOL_VERSION = 1
MAX_PAYLOAD = 16 * 1024 * 1024

def _fingerprint(data): return hashlib.sha256(data).hexdigest()

def _convert(item, root, xsl_path):
    formula_id = str(item.get("formulaId", "")); relative = Path(str(item.get("olePath", "")))
    expected = str(item.get("fingerprint", "")).lower()
    result = {"formulaId": formula_id, "status": "corrupt", "risk": "blocked",
              "mtefVersion": None, "fingerprint": expected, "mathMl": "", "omml": "",
              "warnings": [], "errors": [], "reasonCode": "SIDECAR_UNKNOWN"}
    try:
        path = (root / relative).resolve()
        if root not in path.parents or not path.is_file(): raise ValueError("OLE path is outside the operation directory")
        if path.stat().st_size <= 0 or path.stat().st_size > MAX_PAYLOAD: raise ValueError("OLE payload size is invalid")
        data = path.read_bytes(); actual = _fingerprint(data); result["fingerprint"] = actual
        if actual != expected: raise ValueError("OLE fingerprint mismatch")
        formula = FormulaInfo(formula_id=formula_id, ole_name=path.name, part_name="", rels_path="",
                              relationship_id="", prog_id="Equation.DSMT4", ole_data=data)
        if not extract_mtef(formula):
            result.update(reasonCode="MTEF_EXTRACT_FAILED", errors=[formula.error_message or "MTEF extraction failed"]); return result
        formula.mtef_version = detect_mtef_version(formula.mtef_bytes)
        convert_formula(formula, xsl_path=xsl_path)
        result.update(status="convertible" if formula.status == FormulaStatus.CONVERTED else "unsupported",
                      risk=formula.risk_level.value, mtefVersion=formula.mtef_version,
                      mathMl=formula.mathml or "", omml=formula.omml or "",
                      warnings=formula.quality_warnings,
                      errors=formula.quality_errors or ([formula.error_message] if formula.error_message else []),
                      reasonCode="OK" if formula.status == FormulaStatus.CONVERTED else "CONVERSION_REJECTED")
    except Exception as error:
        result.update(reasonCode="SIDECAR_ITEM_FAILED", errors=[str(error)])
    return result

def main():
    if len(sys.argv) != 3: return 2
    manifest_path = Path(sys.argv[1]).resolve(); output_path = Path(sys.argv[2]).resolve(); root = manifest_path.parent
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    if manifest.get("protocolVersion") != PROTOCOL_VERSION: raise ValueError("Unsupported sidecar protocol version")
    xsl_path = str(manifest.get("officeXsltPath", ""))
    if not Path(xsl_path).is_file(): raise FileNotFoundError("Microsoft Office MML2OMML.XSL was not found")
    workers = max(1, min(4, int(manifest.get("maxWorkers", 4)))); items = list(manifest.get("items", []))
    with ThreadPoolExecutor(max_workers=workers) as executor:
        results = list(executor.map(lambda value: _convert(value, root, xsl_path), items))
    payload = {"protocolVersion": PROTOCOL_VERSION, "operationId": str(manifest.get("operationId", "")), "items": results}
    temporary = output_path.with_suffix(output_path.suffix + ".tmp")
    temporary.write_text(json.dumps(payload, ensure_ascii=True), encoding="utf-8"); os.replace(temporary, output_path)
    return 0

if __name__ == "__main__": raise SystemExit(main())
