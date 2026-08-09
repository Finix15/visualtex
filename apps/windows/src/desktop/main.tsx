import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import "mathlive/static.css";
import "../styles.css";
import "../styles-macos-main.css";
import "../styles-editor-parity.css";
import "../styles-latest-macos-ui.css";
import "../styles-windows-shared-latest.css";
import "../math/customSymbolRendering";
import "../ocr/windowsQuickOcrRuntime";
import { configureOcrTransport } from "../ocr/ocrService";
import { desktopOcrTransport } from "../ocr/ocrTransport";
import { DesktopApp } from "./DesktopApp";
import { applyDocumentTheme, readSynchronizedTheme } from "../themeSync";

configureOcrTransport(desktopOcrTransport);
applyDocumentTheme(readSynchronizedTheme());

createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <DesktopApp />
  </StrictMode>,
);
