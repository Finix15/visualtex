import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

const quickOcr = await readFile("src/ocr/quickOcr.ts", "utf8");
const app = await readFile("src/App.tsx", "utf8");
const nativeBridge = await readFile("src-tauri/src/windows_quick_ocr.rs", "utf8");
const silentNative = await readFile("src-tauri/src/windows_silent_ocr_hotkey.rs", "utf8");
const tauriLib = await readFile("src-tauri/src/lib.rs", "utf8");
const shortcutRuntime = await readFile("src/ocr/windowsQuickOcrRuntime.ts", "utf8");

function functionBody(source, signature, nextMarker) {
  const start = source.indexOf(signature);
  assert.notEqual(start, -1, `Missing ${signature}`);
  const end = source.indexOf(nextMarker, start + signature.length);
  assert.notEqual(end, -1, `Missing end marker ${nextMarker}`);
  return source.slice(start, end);
}

const immediate = functionBody(
  quickOcr,
  "export async function captureQuickOcrScreenshot()",
  "/**\n * The alternate workflow",
);
assert.ok(!immediate.includes("minimizeForOcrCapture"));
assert.ok(immediate.includes("captureWindowsClipboardImage(true)"));

const systemScreenshot = functionBody(
  quickOcr,
  "export async function waitForQuickOcrSystemScreenshot()",
  "/** Register the Windows global Ctrl+Alt+O bridge",
);
assert.ok(!systemScreenshot.includes("minimizeForOcrCapture"));
assert.ok(systemScreenshot.includes("captureWindowsClipboardImage(false)"));

const restoreWindow = functionBody(
  quickOcr,
  "export async function restoreQuickOcrWindow()",
  "export async function writeSilentOcrClipboardText",
);
for (const operation of ["current.show()", "current.unminimize()", "current.setFocus()"]) {
  assert.ok(restoreWindow.includes(operation), `Quick OCR restore is missing ${operation}`);
}

const quickHandler = functionBody(
  app,
  "const handleQuickOcr = async () =>",
  "const handleQuickOcrCaptureModeChange =",
);
assert.ok(quickHandler.includes("finally"));
assert.ok(quickHandler.includes("await restoreQuickOcrWindow()"));
assert.ok(!app.includes("handleSilentOcrShortcut"));

// Interactive quick OCR is minimized natively so the frontend cannot swallow
// a failed/focused-window minimization before Windows Snipping Tool opens.
assert.ok(nativeBridge.includes('get_webview_window("main")'));
assert.ok(nativeBridge.includes(".minimize()"));
assert.ok(nativeBridge.includes("capture_windows_quick_ocr_bytes"));

// Silent OCR is fully native and must not depend on a React/WebView event.
for (const token of [
  "run_silent_ocr",
  "capture_windows_quick_ocr_bytes",
  "format_silent_ocr_latex",
  "OcrImageRequest",
  "write_clipboard_text",
  "silent-ocr.json",
]) {
  assert.ok(silentNative.includes(token), `Native silent OCR is missing ${token}`);
}
assert.ok(!silentNative.includes("visualtex-silent-ocr-global"));
assert.ok(!shortcutRuntime.includes("visualtex-silent-ocr"));
assert.ok(tauriLib.includes("windows_silent_ocr_hotkey::configure"));
assert.ok(tauriLib.includes("windows_quick_ocr::capture_windows_quick_ocr"));

console.log("Windows Quick OCR and native silent OCR migration regression passed");
