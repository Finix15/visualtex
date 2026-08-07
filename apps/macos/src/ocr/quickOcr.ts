import { invoke } from "@tauri-apps/api/core";
import type { LatexCodeFormat } from "../types/formula";
import type { OcrModelName } from "./ocrService";

export const SILENT_OCR_STORAGE_KEY = "visualtex.silent-ocr.enabled";
export const SILENT_OCR_SHORTCUT = "⌘⇧O";

export interface QuickOcrCapture {
  dataBase64: string;
  extension: string;
}

export async function captureQuickOcrScreenshot() {
  return invoke<QuickOcrCapture | null>("capture_quick_ocr_screenshot");
}

export async function configureSilentOcr(
  enabled: boolean,
  model: OcrModelName,
  copyFormat: LatexCodeFormat,
) {
  await invoke("configure_silent_ocr", { enabled, model, copyFormat });
}

export function quickOcrCaptureToFile(capture: QuickOcrCapture) {
  const binary = atob(capture.dataBase64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return new File([bytes], `VisualTeX-Quick-OCR.${capture.extension}`, {
    type: capture.extension === "png" ? "image/png" : "application/octet-stream",
  });
}
