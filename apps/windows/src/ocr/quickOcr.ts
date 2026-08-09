import { invoke } from "@tauri-apps/api/core";
import { getCurrentWindow } from "@tauri-apps/api/window";
import type { LatexCodeFormat } from "../types/formula";
import type { OcrModelName } from "./ocrService";

export const SILENT_OCR_STORAGE_KEY = "visualtex.silent-ocr.enabled";
export const SILENT_OCR_SHORTCUT = "Ctrl+Alt+O";
export const QUICK_OCR_CAPTURE_MODE_STORAGE_KEY =
  "visualtex.quick-ocr.capture-mode";

export type QuickOcrCaptureMode = "immediate" | "system-screenshot";

export function isQuickOcrCaptureMode(
  value: unknown,
): value is QuickOcrCaptureMode {
  return value === "immediate" || value === "system-screenshot";
}

export interface QuickOcrCapture {
  dataBase64: string;
  extension: string;
}

function hasTauriRuntime() {
  return Boolean(
    (window as Window & { __TAURI_INTERNALS__?: { metadata?: unknown } })
      .__TAURI_INTERNALS__?.metadata,
  );
}

async function captureWindowsClipboardImage(launchCapture: boolean) {
  if (!hasTauriRuntime()) {
    throw new Error("Windows Quick OCR is available in the desktop app only.");
  }
  return invoke<QuickOcrCapture | null>("capture_windows_quick_ocr", {
    launchCapture,
    timeoutMs: 60_000,
  });
}

async function minimizeForOcrCapture() {
  try {
    await getCurrentWindow().minimize();
    // Give DWM/WebView2 one frame to leave the capture surface before the
    // Snipping Tool is opened. Without this delay the VisualTeX window can
    // still appear in the first captured frame on slower systems.
    await new Promise((resolve) => window.setTimeout(resolve, 160));
  } catch {
    // The native capture bridge still works if the host refuses minimization.
  }
}

export async function restoreQuickOcrWindow() {
  if (!hasTauriRuntime()) return;
  try {
    const current = getCurrentWindow();
    await current.show();
    await current.unminimize();
    await current.setFocus();
  } catch {
    // OCR insertion can still finish even if Windows refuses focus restoration.
  }
}

export async function writeSilentOcrClipboardText(text: string) {
  if (!hasTauriRuntime()) {
    throw new Error("Windows silent OCR clipboard access is unavailable.");
  }
  await invoke("write_windows_ocr_clipboard_text", { text });
}

/**
 * Windows "immediate" capture opens the native Snipping Tool region selector
 * immediately. This keeps capture precise and avoids a second, app-specific
 * desktop-capture implementation.
 */
export async function captureQuickOcrScreenshot() {
  await minimizeForOcrCapture();
  return captureWindowsClipboardImage(true);
}

/**
 * The alternate workflow intentionally does not open Snipping Tool. VisualTeX
 * minimizes so the user can navigate to any page and invoke their preferred
 * Windows screenshot shortcut; the next clipboard image is consumed by OCR.
 */
export async function waitForQuickOcrSystemScreenshot() {
  await minimizeForOcrCapture();
  return captureWindowsClipboardImage(false);
}

/** Register the Windows global Ctrl+Alt+O bridge while OCR stays in React. */
export async function configureSilentOcr(
  enabled: boolean,
  model: OcrModelName,
  copyFormat: LatexCodeFormat,
) {
  if (!hasTauriRuntime()) return;
  await invoke("configure_silent_ocr", { enabled, model, copyFormat });
}

export function quickOcrCaptureToFile(capture: QuickOcrCapture) {
  const binary = atob(capture.dataBase64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  const extension = capture.extension || "png";
  return new File([bytes], `VisualTeX-Quick-OCR.${extension}`, {
    type:
      extension === "jpg" || extension === "jpeg"
        ? "image/jpeg"
        : extension === "webp"
          ? "image/webp"
          : "image/png",
  });
}
