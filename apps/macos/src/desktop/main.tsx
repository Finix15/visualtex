import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import "mathlive/static.css";
import "../styles.css";
import "../styles-editor-parity.css";
import { configureOcrTransport } from "../ocr/ocrService";
import { desktopOcrTransport } from "../ocr/ocrTransport";
import { OfficeDialogApp } from "../office/dialog/OfficeDialogApp";
import { OfficeDocumentImportApp } from "../office/documentImport/OfficeDocumentImportApp";
import { DesktopApp } from "./DesktopApp";
import { applyDocumentTheme, readSynchronizedTheme } from "../themeSync";
import { VisualTexErrorBoundary } from "../runtime/VisualTexErrorBoundary";

configureOcrTransport(desktopOcrTransport);

const root = document.getElementById("root");
if (!root) throw new Error("Missing VisualTeX application root element.");

const view = new URLSearchParams(window.location.search).get("view");
const officeFormulaView = view === "office-formula";
const officeDocumentImportView = view === "office-document-import";
if (officeFormulaView || officeDocumentImportView) {
  document.body.classList.add("office-dialog-page");
  applyDocumentTheme(readSynchronizedTheme());
}

createRoot(root).render(
  <StrictMode>
    <VisualTexErrorBoundary>
      {officeFormulaView ? (
        <OfficeDialogApp />
      ) : officeDocumentImportView ? (
        <OfficeDocumentImportApp />
      ) : (
        <DesktopApp />
      )}
    </VisualTexErrorBoundary>
  </StrictMode>,
);
