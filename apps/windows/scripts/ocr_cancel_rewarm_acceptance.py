from __future__ import annotations

import argparse
import json
import os
import tempfile
import threading
import time
from pathlib import Path

import requests
import urllib3
from PIL import Image, ImageDraw, ImageFont

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


def read_increment(path: Path, offset: int) -> bytes:
    if not path.exists():
        return b""
    with path.open("rb") as handle:
        handle.seek(offset)
        return handle.read()


def wait_for_text(path: Path, offset: int, needle: str, timeout: float) -> tuple[bool, bytes]:
    deadline = time.monotonic() + timeout
    data = b""
    while time.monotonic() < deadline:
        data = read_increment(path, offset)
        if needle in data.decode("utf-8", errors="replace"):
            return True, data
        time.sleep(0.05)
    return False, data


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--model", default="PP-FormulaNet_plus-M")
    args = parser.parse_args()

    app_data = Path(os.environ["APPDATA"]) / "com.visualtex.studio"
    install = json.loads((app_data / "office" / "install.json").read_text(encoding="utf-8"))
    token = install["installToken"]
    port = int(install.get("port", 43127))
    base = f"https://127.0.0.1:{port}/api/v1/ocr"
    auth = {"x-visualtex-install-token": token}

    session = requests.Session()
    session.verify = False
    health = session.get(f"https://127.0.0.1:{port}/health", timeout=5)
    health.raise_for_status()

    log_path = app_data / "ocr-runtime" / "logs" / "worker.log"
    before = log_path.stat().st_size if log_path.exists() else 0

    restart = session.post(f"{base}/restart", headers=auth, timeout=10)
    restart.raise_for_status()

    with tempfile.TemporaryDirectory(prefix="visualtex-ocr-cancel-") as temp_dir:
        image_path = Path(temp_dir) / "formula.png"
        image = Image.new("RGB", (720, 180), "white")
        draw = ImageDraw.Draw(image)
        try:
            font = ImageFont.truetype(r"C:\Windows\Fonts\cambria.ttc", 72)
        except OSError:
            font = ImageFont.load_default()
        draw.text((35, 45), "x² + y² = 1", fill="black", font=font)
        image.save(image_path)
        image_bytes = image_path.read_bytes()

        first_result: dict[str, object] = {}

        def recognize_once() -> None:
            started = time.perf_counter()
            try:
                response = session.post(
                    f"{base}/recognize",
                    headers={
                        **auth,
                        "content-type": "application/octet-stream",
                        "x-visualtex-ocr-model": args.model,
                        "x-visualtex-ocr-extension": "png",
                    },
                    data=image_bytes,
                    timeout=120,
                )
                first_result["status"] = response.status_code
                first_result["body"] = response.text
            except Exception as error:  # noqa: BLE001 - acceptance captures transport failures
                first_result["error"] = repr(error)
            first_result["elapsed_ms"] = round((time.perf_counter() - started) * 1000)

        thread = threading.Thread(target=recognize_once, daemon=True)
        thread.start()

        loading_marker = f"Loading {args.model} directly from local model_dir"
        loading_seen, _ = wait_for_text(log_path, before, loading_marker, timeout=15)
        if not loading_seen:
            raise AssertionError("Recognition never entered local model loading before cancellation")

        cancel_started = time.perf_counter()
        cancel = session.post(f"{base}/cancel", headers=auth, timeout=10)
        cancel.raise_for_status()
        cancel_ms = round((time.perf_counter() - cancel_started) * 1000)
        thread.join(timeout=20)
        if thread.is_alive():
            raise AssertionError("Cancelled recognition request did not return")

        cancel_offset = log_path.stat().st_size if log_path.exists() else before
        warmed_marker = f"Warmed {args.model} on cpu in "
        recovered, recovery_data = wait_for_text(
            log_path,
            cancel_offset,
            warmed_marker,
            timeout=20,
        )
        if not recovered:
            tail = recovery_data.decode("utf-8", errors="replace")[-2000:]
            raise AssertionError(f"Worker did not automatically rewarm after cancellation. log tail={tail}")

        warm_started = time.perf_counter()
        warm = session.post(
            f"{base}/warmup",
            headers={**auth, "x-visualtex-ocr-model": args.model},
            timeout=20,
        )
        warm.raise_for_status()
        warm_reuse_ms = round((time.perf_counter() - warm_started) * 1000)
        if warm_reuse_ms > 1500:
            raise AssertionError(
                f"Warmup after recovery took {warm_reuse_ms} ms; the recovered model was not reused"
            )

        after_recovery_offset = log_path.stat().st_size
        second_started = time.perf_counter()
        second = session.post(
            f"{base}/recognize",
            headers={
                **auth,
                "content-type": "application/octet-stream",
                "x-visualtex-ocr-model": args.model,
                "x-visualtex-ocr-extension": "png",
            },
            data=image_bytes,
            timeout=60,
        )
        second_ms = round((time.perf_counter() - second_started) * 1000)
        second_delta = read_increment(log_path, after_recovery_offset).decode(
            "utf-8", errors="replace"
        )
        if loading_marker in second_delta:
            raise AssertionError("Second recognition reloaded the model after cancellation recovery")

    print(
        json.dumps(
            {
                "model": args.model,
                "cancelMs": cancel_ms,
                "firstRecognition": first_result,
                "recovered": recovered,
                "warmReuseMs": warm_reuse_ms,
                "secondStatus": second.status_code,
                "secondElapsedMs": second_ms,
                "secondReloadedModel": False,
            },
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
