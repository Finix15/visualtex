"""Offline acceptance probe for the packaged sidecar runtime."""
import argparse, hashlib, json, subprocess, tempfile, zipfile
from pathlib import Path

def main():
    parser = argparse.ArgumentParser(); parser.add_argument("docx", type=Path); parser.add_argument("runtime", type=Path); parser.add_argument("xsl", type=Path); parser.add_argument("--expect", type=int, required=True)
    args = parser.parse_args()
    with tempfile.TemporaryDirectory(prefix="visualtex-mathtype-sidecar-") as temporary:
        root = Path(temporary); items = []
        with zipfile.ZipFile(args.docx) as archive:
            names = sorted(name for name in archive.namelist() if name.startswith("word/embeddings/") and name.endswith(".bin"))
            if len(names) != args.expect: raise AssertionError(f"expected {args.expect} embeddings, found {len(names)}")
            for index, name in enumerate(names):
                data = archive.read(name); output = root / f"formula-{index:04d}.bin"; output.write_bytes(data)
                items.append({"formulaId": f"F{index + 1:04d}", "olePath": output.name, "fingerprint": hashlib.sha256(data).hexdigest()})
        operation_id = "fixture-smoke"; request = root / "request.json"; response = root / "response.json"
        request.write_text(json.dumps({"protocolVersion": 1, "operationId": operation_id, "officeXsltPath": str(args.xsl), "maxWorkers": 4, "items": items}), encoding="utf-8")
        completed = subprocess.run([args.runtime / "python.exe", "-I", "-X", "utf8", args.runtime / "worker.py", request, response], timeout=180, check=True)
        payload = json.loads(response.read_text(encoding="utf-8")); results = payload["items"]
        if payload["operationId"] != operation_id or len(results) != args.expect: raise AssertionError("sidecar response mismatch")
        converted = sum(item["status"] == "convertible" for item in results)
        failed = len(results) - converted
        print(json.dumps({"total": len(results), "converted": converted, "failed": failed}))
        return completed.returncode

if __name__ == "__main__": raise SystemExit(main())
