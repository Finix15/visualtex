from __future__ import annotations

import importlib.util
import os
import sys
import tempfile
import types
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
WORKER_PATH = ROOT / "src-tauri" / "ocr" / "worker.py"
LIB_PATH = ROOT / "src-tauri" / "src" / "lib.rs"


def load_worker():
    spec = importlib.util.spec_from_file_location("visualtex_ocr_worker_regression", WORKER_PATH)
    if spec is None or spec.loader is None:
        raise AssertionError("Unable to load OCR worker module")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def main() -> None:
    worker = load_worker()
    original_cache = os.environ.get("PADDLE_PDX_CACHE_HOME")
    original_home = os.environ.get("HOME")
    original_userprofile = os.environ.get("USERPROFILE")
    original_paddleocr = sys.modules.get("paddleocr")

    try:
        with tempfile.TemporaryDirectory(prefix="visualtex-ocr-warmup-") as temp_root:
            cache_root = Path(temp_root)
            model_name = "PP-FormulaNet_plus-M"
            model_dir = cache_root / "official_models" / model_name
            model_dir.mkdir(parents=True)
            for filename in ("inference.json", "inference.pdiparams", "inference.yml"):
                (model_dir / filename).write_bytes(b"test")

            isolated_home = cache_root / "isolated-home"
            isolated_home.mkdir()
            os.environ["PADDLE_PDX_CACHE_HOME"] = str(cache_root)
            os.environ["HOME"] = str(isolated_home)
            os.environ["USERPROFILE"] = str(isolated_home)
            calls: list[dict[str, object]] = []

            fake_paddleocr = types.ModuleType("paddleocr")

            class FakeFormulaRecognition:
                def __init__(self, **kwargs):
                    calls.append(kwargs)

            fake_paddleocr.FormulaRecognition = FakeFormulaRecognition
            sys.modules["paddleocr"] = fake_paddleocr

            worker._CURRENT_MODEL = None
            worker._CURRENT_MODEL_NAME = None
            worker._CURRENT_DEVICE = None
            first = worker._load_model(model_name, "cpu")
            second = worker._load_model(model_name, "cpu")

            assert first is second, "A warmed model must be reused in the same worker"
            assert len(calls) == 1, "Repeated warmup must not reconstruct the same model"
            assert calls[0].get("model_name") == model_name
            assert calls[0].get("device") == "cpu"
            assert Path(str(calls[0].get("model_dir"))).resolve() == model_dir.resolve(), (
                "A complete local cache must be passed as model_dir so PaddleX does not "
                "perform model host checks or metadata downloads"
            )

            worker._CURRENT_MODEL = None
            worker._CURRENT_MODEL_NAME = None
            worker._CURRENT_DEVICE = None
            calls.clear()
            (model_dir / "inference.yml").unlink()
            try:
                worker._load_model(model_name, "cpu")
            except RuntimeError as error:
                assert "verified local model" in str(error).lower()
            else:
                raise AssertionError(
                    "An incomplete model directory must be rejected instead of allowing PaddleX to download implicitly"
                )
            assert calls == [], (
                "An incomplete model directory must fail before PaddleOCR is constructed"
            )
    finally:
        if original_cache is None:
            os.environ.pop("PADDLE_PDX_CACHE_HOME", None)
        else:
            os.environ["PADDLE_PDX_CACHE_HOME"] = original_cache
        if original_home is None:
            os.environ.pop("HOME", None)
        else:
            os.environ["HOME"] = original_home
        if original_userprofile is None:
            os.environ.pop("USERPROFILE", None)
        else:
            os.environ["USERPROFILE"] = original_userprofile
        if original_paddleocr is None:
            sys.modules.pop("paddleocr", None)
        else:
            sys.modules["paddleocr"] = original_paddleocr

    rust_source = LIB_PATH.read_text(encoding="utf-8")
    required_fragments = (
        "schedule_startup_warmup",
        "schedule_worker_rewarm(",
        "Unable to restore OCR prewarm after cancellation",
        "office_ocr_state.schedule_startup_warmup",
        "DEFAULT_OCR_MODEL",
    )
    for fragment in required_fragments:
        assert fragment in rust_source, f"Missing OCR warmup lifecycle guard: {fragment}"

    cancel_start = rust_source.index("pub(crate) fn cancel(&self, app: &AppHandle)")
    cancel_end = rust_source.index("pub(crate) async fn restart", cancel_start)
    cancel_source = rust_source[cancel_start:cancel_end]
    assert "terminate_worker_process" in cancel_source
    assert "schedule_worker_rewarm" in cancel_source

    print("VisualTeX OCR local-model warmup and cancel recovery regression passed")


if __name__ == "__main__":
    main()
