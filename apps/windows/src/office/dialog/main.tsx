import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import "mathlive/static.css";
import "../../styles.css";
import "../../styles-macos-main.css";
import { configureOcrTransport } from "../../ocr/ocrService";
import { officeOcrTransport } from "../api/ocrHttpTransport";
import { OfficeDialogApp } from "./OfficeDialogApp";
import { DocumentImportApp } from "../documentImport/DocumentImportApp";
import { applyDocumentTheme, readSynchronizedTheme } from "../../themeSync";
import {
  normalizeEditorLayout,
  useEditorStore,
} from "../../stores/editorStore";

configureOcrTransport(officeOcrTransport);
applyDocumentTheme(readSynchronizedTheme());
useEditorStore.getState().setEditorLayout(
  normalizeEditorLayout(
    document
      .querySelector<HTMLMetaElement>('meta[name="visualtex-editor-layout"]')
      ?.content,
  ),
);

function mount() {
  const root = document.getElementById("root");
  if (!root) throw new Error("Missing Office Dialog root element.");
  const runtime = new URLSearchParams(window.location.search).get("runtime");
  createRoot(root).render(
    <StrictMode>
      {runtime === "vsto-bulk-import" ? <DocumentImportApp /> : <OfficeDialogApp />}
    </StrictMode>,
  );
}

// Windows Office integration is now exclusively driven by the native Ribbon
// COM add-ins. The dialog is hosted by the VisualTeX companion service and no
// longer waits for, imports or executes the Office.js runtime.
mount();
