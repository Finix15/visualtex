import { useEffect, useMemo, useRef, useState } from "react";
import { AlertCircle, LoaderCircle, X } from "lucide-react";
import {
  cancelMacosDocumentImport,
  closeMacosDocumentImportWindow,
  commitMacosDocumentImport,
  focusMacosDocumentImportTarget,
  getMacosDocumentImportRequest,
  reportMacosLatexRedrawStage,
  resolveMacosLatexRedrawFontSizes,
  restoreMacosDocumentImportWindow,
} from "../documentImport/documentImportClient";
import { documentImportErrorMessage } from "../documentImport/documentImportErrors";
import { findWindowsWordLatexRedrawSpans } from "./wordLatexRedrawParser";
import { prepareWindowsStyleWordLatexRedrawItems } from "./wordLatexRedrawRenderer";

export function WordLatexRedrawApp() {
  const sessionId = useMemo(
    () => new URLSearchParams(window.location.search).get("sessionId") ?? "",
    [],
  );
  const startedRef = useRef(false);
  const [status, setStatus] = useState("Preparing Word LaTeX redraw…");
  const [error, setError] = useState("");
  const [closing, setClosing] = useState(false);

  useEffect(() => {
    if (!sessionId || startedRef.current) return;
    startedRef.current = true;

    void (async () => {
      const workflowStarted = performance.now();
      const reportStage = (stage: string, itemCount: number) =>
        reportMacosLatexRedrawStage(
          sessionId,
          stage,
          performance.now() - workflowStarted,
          itemCount,
        ).catch(() => undefined);
      try {
        const request = await getMacosDocumentImportRequest(sessionId);
        await reportStage("latex-redraw-request-ready", 0);
        if (request.operation !== "latexRedraw") {
          throw new Error("The current Office session is not a Word LaTeX redraw request.");
        }
        const source = request.source ?? "";
        const outputKind = request.outputKind;
        if (!source) throw new Error("The Word LaTeX redraw source is empty.");
        if (outputKind !== "omml" && outputKind !== "image") {
          throw new Error("The Word LaTeX redraw output format is invalid.");
        }

        const spans = findWindowsWordLatexRedrawSpans(source);
        await reportStage("latex-redraw-parse-complete", spans.length);
        if (!spans.length) {
          throw new Error(
            request.redrawScope === "document"
              ? "No complete LaTeX formulas were found in the Word document."
              : "No complete LaTeX formulas were found in the selected Word text.",
          );
        }

        await focusMacosDocumentImportTarget("latexRedraw");
        setStatus(`Reading Word font sizes for ${spans.length} formulas…`);
        const fontSizes = await resolveMacosLatexRedrawFontSizes(
          sessionId,
          spans.map((span) => ({
            sourceStart: span.start,
            sourceEnd: span.end,
            sourceText: span.sourceText,
            displayMode: span.displayMode,
          })),
        );
        if (fontSizes.length !== spans.length) {
          throw new Error("Word returned an incomplete LaTeX redraw font plan.");
        }
        await reportStage("latex-redraw-font-preflight-complete", spans.length);
        const targets = spans.map((span, index) => ({
          ...span,
          fontSizePt: fontSizes[index],
        }));
        setStatus(`Rendering 0/${targets.length} formulas…`);
        const items = await prepareWindowsStyleWordLatexRedrawItems(
          targets,
          outputKind,
          (current, total) => setStatus(`Rendering ${current}/${total} formulas…`),
        );
        await reportStage("latex-redraw-render-complete", items.length);
        setStatus(`Writing ${items.length} formulas back to Word…`);
        await commitMacosDocumentImport(sessionId, { outputKind, items });
        await reportStage("latex-redraw-commit-complete", items.length);
        await closeMacosDocumentImportWindow();
      } catch (reason) {
        setError(documentImportErrorMessage(reason, "Word LaTeX redraw failed."));
        setStatus("");
        await restoreMacosDocumentImportWindow().catch(() => undefined);
      }
    })();
  }, [sessionId]);

  const close = async () => {
    if (closing) return;
    setClosing(true);
    try {
      if (sessionId) await cancelMacosDocumentImport(sessionId).catch(() => undefined);
      await closeMacosDocumentImportWindow();
    } finally {
      setClosing(false);
    }
  };

  if (!error) {
    return (
      <main className="document-import-state" aria-live="polite">
        <LoaderCircle className="is-spinning" />
        <span>{status}</span>
      </main>
    );
  }

  return (
    <main className="document-import-state is-error" role="alert">
      <AlertCircle />
      <strong>Word LaTeX redraw failed</strong>
      <p>{error}</p>
      <button type="button" onClick={() => void close()} disabled={closing}>
        <X size={16} />
        Close
      </button>
    </main>
  );
}
