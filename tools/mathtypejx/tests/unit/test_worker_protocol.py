import hashlib
import json
import uuid
from pathlib import Path

import pytest

from mathtypejx.mtef import mtef_to_mathml
from mathtypejx.worker import finalize, scan


def _request(job_root: Path, source: Path) -> str:
    job_id = str(uuid.uuid4())
    job_root.mkdir()
    job_root.joinpath("request.json").write_text(json.dumps({
        "protocolVersion": 1,
        "jobId": job_id,
        "inputPath": str(source.resolve()),
        "outputPath": str(source.with_name("output.docx").resolve()),
        "inputSha256": hashlib.sha256(source.read_bytes()).hexdigest(),
        "status": "created",
        "createdAtEpochSeconds": 1,
        "updatedAtEpochSeconds": 1,
        "error": None,
    }), encoding="utf-8")
    return job_id


def test_zero_formula_scan_and_finalize(sample_docx_no_formulas, temp_dir):
    root = temp_dir / "job"
    job_id = _request(root, sample_docx_no_formulas)
    scan(job_id, root)
    report = json.loads(root.joinpath("scan-report.json").read_text())
    assert report["detected"] == 0
    assert report["batchCount"] == 0

    finalize(job_id, root)
    conversion = json.loads(root.joinpath("conversion-report.json").read_text())
    assert conversion["detected"] == 0
    assert conversion["sourceUnmodified"] is True
    assert conversion["packageValid"] is True
    assert root.joinpath("candidate.docx").is_file()


@pytest.mark.parametrize("fixture_name", ["oleObject1.bin", "oleObject2.bin", "oleObject3.bin"])
def test_bundled_parser_converts_all_mtef_v5_contract_fixtures(fixture_name):
    fixture = Path(__file__).parents[1] / "fixtures" / fixture_name
    mathml = mtef_to_mathml(fixture.read_bytes())
    assert "<math" in mathml
    assert "Math/MathML" in mathml


def test_finalize_rejects_missing_formula_accounting(sample_docx_no_formulas, temp_dir):
    root = temp_dir / "job"
    job_id = _request(root, sample_docx_no_formulas)
    scan(job_id, root)
    report_path = root / "scan-report.json"
    report = json.loads(report_path.read_text())
    report["batchCount"] = 1
    report_path.write_text(json.dumps(report), encoding="utf-8")
    with pytest.raises(FileNotFoundError):
        finalize(job_id, root)
