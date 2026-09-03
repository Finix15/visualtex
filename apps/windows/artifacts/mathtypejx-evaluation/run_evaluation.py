import argparse
import hashlib
import json
import time
import zipfile
from pathlib import Path

from lxml import etree
from mathtypejx import convert_mathtype_to_omml


NS = {
    "o": "urn:schemas-microsoft-com:office:office",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def canonical_hash(xml_text: str | None) -> str | None:
    if not xml_text:
        return None
    node = etree.fromstring(xml_text.encode("utf-8"))
    return sha256(etree.tostring(node, method="c14n"))


def package_metrics(path: Path) -> dict:
    metrics = {
        "sha256": sha256(path.read_bytes()),
        "zip_test": None,
        "entries": 0,
        "ole_equation_dsmt4": 0,
        "omml_objects": 0,
        "paragraphs": 0,
        "tables": 0,
        "text_sha256": None,
        "media": {},
        "embeddings": {},
    }
    texts = []
    with zipfile.ZipFile(path) as zf:
        metrics["zip_test"] = zf.testzip()
        metrics["entries"] = len(zf.infolist())
        for name in zf.namelist():
            data = zf.read(name)
            if name.startswith("word/media/"):
                metrics["media"][name] = sha256(data)
            if name.startswith("word/embeddings/"):
                metrics["embeddings"][Path(name).name] = sha256(data)
            if name.startswith("word/") and name.endswith(".xml"):
                try:
                    root = etree.fromstring(data)
                except etree.XMLSyntaxError:
                    continue
                metrics["ole_equation_dsmt4"] += len(root.xpath(".//o:OLEObject[@ProgID='Equation.DSMT4']", namespaces=NS))
                metrics["omml_objects"] += len(root.xpath(".//m:oMath", namespaces=NS))
                metrics["paragraphs"] += len(root.xpath(".//w:p", namespaces=NS))
                metrics["tables"] += len(root.xpath(".//w:tbl", namespaces=NS))
                texts.extend(root.xpath(".//w:t/text()", namespaces=NS))
    metrics["text_sha256"] = sha256("\u241e".join(texts).encode("utf-8"))
    return metrics


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--report", type=Path, required=True)
    parser.add_argument("--parallel", action="store_true")
    args = parser.parse_args()

    before = package_metrics(args.input)
    progress_log = []
    started = time.perf_counter()
    report = convert_mathtype_to_omml(
        str(args.input),
        str(args.output),
        remove_edit_info=False,
        parallel=args.parallel,
        max_workers=4,
        progress_callback=lambda phase, current, total: progress_log.append(
            {"phase": phase, "current": current, "total": total, "elapsed_seconds": round(time.perf_counter() - started, 6)}
        ),
    )
    elapsed = time.perf_counter() - started
    after = package_metrics(args.output)

    payload = report.to_dict()
    payload.update({
        "elapsed_seconds": elapsed,
        "parallel": args.parallel,
        "max_workers": 4 if args.parallel else 1,
        "remove_edit_info": False,
        "input_metrics": before,
        "output_metrics": after,
        "progress": progress_log,
    })
    for item, formula in zip(payload["formulas"], report.formulas):
        item["ole_sha256"] = sha256(formula.ole_data) if formula.ole_data else None
        item["mtef_sha256"] = sha256(formula.mtef_bytes) if formula.mtef_bytes else None
        item["mathml_c14n_sha256"] = canonical_hash(formula.mathml)
        item["omml_c14n_sha256"] = canonical_hash(formula.omml)
        item["mathml"] = formula.mathml
        item["omml"] = formula.omml

    args.report.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({
        "total": payload["total_ole_objects"],
        "succeeded": payload["succeeded"],
        "failed": payload["failed"],
        "skipped": payload["skipped"],
        "elapsed_seconds": elapsed,
        "output_ole": after["ole_equation_dsmt4"],
        "output_omml": after["omml_objects"],
    }, ensure_ascii=False))


if __name__ == "__main__":
    main()
