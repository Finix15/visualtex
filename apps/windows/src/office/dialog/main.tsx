import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import "mathlive/static.css";
import "../../styles.css";
import "../../styles-macos-main.css";
import { configureOcrTransport } from "../../ocr/ocrService";
import { officeOcrTransport } from "../api/ocrHttpTransport";
import { OfficeDialogApp } from "./OfficeDialogApp";

configureOcrTransport(officeOcrTransport);

function mount() {
  const root = document.getElementById("root");
  if (!root) throw new Error("Missing Office Dialog root element.");
  createRoot(root).render(
    <StrictMode>
      <OfficeDialogApp />
    </StrictMode>,
  );
}

// Windows Office integration is now exclusively driven by the native Ribbon
// COM add-ins. The dialog is hosted by the VisualTeX companion service and no
// longer waits for, imports or executes the Office.js runtime.
mount();
