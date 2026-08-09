import { listen } from "@tauri-apps/api/event";
import { safeStorage } from "../runtime/safeStorage";

const SILENT_OCR_ENABLED_KEY = "visualtex.silent-ocr.enabled";
let installed = false;

function silentEnabled() {
  const value = safeStorage.getItem(SILENT_OCR_ENABLED_KEY);
  return value === "true" || value === "1";
}

function dispatch(name: "visualtex-quick-ocr" | "visualtex-silent-ocr") {
  window.dispatchEvent(new CustomEvent(name));
}

function hasTauriRuntime() {
  return Boolean(
    (window as Window & { __TAURI_INTERNALS__?: { metadata?: unknown } })
      .__TAURI_INTERNALS__?.metadata,
  );
}

/**
 * Windows-only keyboard bridge. OCR execution intentionally stays in the
 * React application so shortcuts, buttons, model selection, progress UI and
 * error handling all use one code path instead of competing DOM/runtime
 * implementations.
 */
export function installWindowsQuickOcrRuntime() {
  if (installed || typeof window === "undefined" || typeof document === "undefined") {
    return;
  }
  installed = true;
  if (hasTauriRuntime()) {
    void listen("visualtex-silent-ocr-global", () => {
      if (silentEnabled()) dispatch("visualtex-silent-ocr");
    }).catch(() => undefined);
  }
  document.addEventListener(
    "keydown",
    (event) => {
      if (event.repeat || event.isComposing) return;
      if (
        event.ctrlKey &&
        event.shiftKey &&
        !event.altKey &&
        !event.metaKey &&
        event.code === "KeyO"
      ) {
        event.preventDefault();
        dispatch("visualtex-quick-ocr");
        return;
      }
      if (
        event.ctrlKey &&
        event.altKey &&
        !event.shiftKey &&
        !event.metaKey &&
        event.code === "KeyO" &&
        silentEnabled()
      ) {
        event.preventDefault();
        dispatch("visualtex-silent-ocr");
      }
    },
    true,
  );
}

installWindowsQuickOcrRuntime();
