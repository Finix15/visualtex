import { Component, createElement, useCallback, useEffect, useRef, useState, type ErrorInfo, type ReactNode } from "react";
import { confirm, open, save } from "@tauri-apps/plugin-dialog";
import { getCurrentWindow } from "@tauri-apps/api/window";
import { useEditorStore } from "../../stores/editorStore";
import { mathMlToOmmlArtifacts } from "../omml/latexToOmml";
import { validateMathMlToOmml } from "../omml/mathMlOmmlValidator";
import { legacyEquationClient } from "./legacyEquationClient";
import { legacyEquationError } from "./legacyEquationErrors";
import { assertFormulaBatch, defaultOutputPath, type FormulaBatch, type LegacyJobView, type OmmlBatch } from "./legacyEquationTypes";

const RECOVERY_KEY = "visualtex.legacy-equation.active-job.v1";
const MAX_PREVIEW_BYTES = 256 * 1024;
const MATHML_ELEMENTS = new Set(["math", "mrow", "mi", "mn", "mo", "mtext", "mfrac", "msqrt", "mroot", "msub", "msup", "msubsup", "mmultiscripts", "mprescripts", "none", "mfenced", "mtable", "mtr", "mtd", "mover", "munder", "munderover", "menclose", "mspace"]);
type UiStage = "idle" | "scanning" | "review" | "converting" | "validating" | "finalizing" | "complete" | "cancelling" | "cancelled" | "failed" | "recovered";

class PreviewBoundary extends Component<{ children: ReactNode; fallback: string }, { failed: boolean }> {
  state = { failed: false };
  static getDerivedStateFromError() { return { failed: true }; }
  componentDidCatch(_error: Error, _info: ErrorInfo) { /* Payload intentionally not logged. */ }
  render() { return this.state.failed ? <span>{this.props.fallback}</span> : this.props.children; }
}

function safeMathMlPreview(source: string): ReactNode {
  const documentObject = new DOMParser().parseFromString(source, "application/xml");
  if (documentObject.querySelector("parsererror") || /<!DOCTYPE|<!ENTITY/i.test(source)) throw new Error("Malformed MathML preview");
  const visit = (node: Node, key: string): ReactNode => {
    if (node.nodeType === Node.TEXT_NODE) return node.textContent;
    if (!(node instanceof Element) || !MATHML_ELEMENTS.has(node.localName)) return null;
    const props: Record<string, string> = { key };
    for (const name of ["display", "mathvariant", "mathcolor", "open", "close", "separators", "notation", "stretchy"]) {
      const value = node.getAttribute(name); if (value !== null && value.length <= 128) props[name] = value;
    }
    return createElement(node.localName, props, Array.from(node.childNodes).map((child, index) => visit(child, `${key}-${index}`)));
  };
  return visit(documentObject.documentElement, "math");
}

function progress(stage: UiStage) {
  return ({ idle: 0, scanning: 15, recovered: 20, review: 35, converting: 55, validating: 75,
    finalizing: 90, complete: 100, cancelling: 50, cancelled: 0, failed: 0 })[stage];
}

export function LegacyEquationConverterApp() {
  const language = useEditorStore((state) => state.language);
  const isEn = language === "en";
  const t = (vi: string, en: string) => isEn ? en : vi;
  const [stage, setStage] = useState<UiStage>("idle");
  const [inputPath, setInputPath] = useState("");
  const [outputPath, setOutputPath] = useState("");
  const [job, setJob] = useState<LegacyJobView | null>(null);
  const [batch, setBatch] = useState<FormulaBatch | null>(null);
  const [page, setPage] = useState(0);
  const [selected, setSelected] = useState<Set<string>>(new Set());
  const [deselected, setDeselected] = useState<Set<string>>(new Set());
  const [error, setError] = useState("");
  const generation = useRef(0);
  const creating = useRef(false);
  const errorRef = useRef<HTMLDivElement>(null);
  const active = ["scanning", "converting", "validating", "finalizing", "cancelling"].includes(stage);

  const fail = useCallback((reason: unknown) => {
    setError(legacyEquationError(reason, isEn ? "en" : "vi"));
    setStage("failed");
    requestAnimationFrame(() => errorRef.current?.focus());
  }, [isEn]);

  const loadBatch = useCallback(async (current: LegacyJobView, index: number, token = generation.current) => {
    const parsed = assertFormulaBatch(await legacyEquationClient.readBatch(current.jobId, index), current.jobId, index);
    if (token !== generation.current) return;
    setBatch(parsed);
    setPage(index);
    setSelected((previous) => {
      const next = new Set(previous);
      for (const formula of parsed.formulas) if (formula.riskLevel === "auto-replace") next.add(formula.formulaId);
      return next;
    });
  }, []);

  useEffect(() => {
    const recovered = localStorage.getItem(RECOVERY_KEY);
    if (!recovered) return;
    const token = ++generation.current;
    legacyEquationClient.get(recovered).then(async (value) => {
      if (token !== generation.current) return;
      setJob(value);
      setInputPath(value.inputPath);
      setOutputPath(value.outputPath);
      setStage(value.status === "complete" ? "complete" : value.status === "failed" ? "failed" : "recovered");
      if (value.status === "awaitingOmml" && value.scanReport?.batchCount) {
        await loadBatch(value, 0, token);
        setStage("review");
      }
    }).catch(() => localStorage.removeItem(RECOVERY_KEY));
  }, [loadBatch]);

  useEffect(() => {
    if (stage !== "scanning" || !job) return;
    const token = generation.current;
    const timer = window.setInterval(() => {
      legacyEquationClient.get(job.jobId).then(async (value) => {
        if (token !== generation.current) return;
        setJob(value);
        if (value.status === "failed") return fail(value.error);
        if (value.status === "awaitingOmml") {
          window.clearInterval(timer);
          if ((value.scanReport?.batchCount ?? 0) > 0) await loadBatch(value, 0, token);
          setStage("review");
        }
      }).catch(fail);
    }, 300);
    return () => window.clearInterval(timer);
  }, [stage, job?.jobId, fail, loadBatch]);

  useEffect(() => {
    const unlisten = getCurrentWindow().onCloseRequested(async (event) => {
      if (!active || !job) return;
      event.preventDefault();
      if (await confirm(t("Chuyển đổi đang chạy. Hủy và đóng cửa sổ?", "Conversion is running. Cancel and close?"))) {
        await legacyEquationClient.cancel(job.jobId).catch(() => undefined);
        localStorage.removeItem(RECOVERY_KEY);
        await getCurrentWindow().destroy();
      }
    });
    return () => { void unlisten.then((dispose) => dispose()); };
  }, [active, job?.jobId, isEn]);

  async function chooseInput() {
    const chosen = await open({ multiple: false, filters: [{ name: "Word", extensions: ["docx"] }] });
    if (typeof chosen !== "string") return;
    setInputPath(chosen);
    setOutputPath(defaultOutputPath(chosen));
  }

  async function chooseOutput() {
    const chosen = await save({ defaultPath: outputPath || defaultOutputPath(inputPath), filters: [{ name: "Word", extensions: ["docx"] }] });
    if (chosen) setOutputPath(chosen);
  }

  async function start() {
    if (creating.current || !inputPath || !outputPath) return;
    creating.current = true;
    const token = ++generation.current;
    setError(""); setStage("scanning"); setBatch(null); setSelected(new Set()); setDeselected(new Set());
    try {
      const value = await legacyEquationClient.create(inputPath, outputPath);
      if (token !== generation.current) return;
      setJob(value); localStorage.setItem(RECOVERY_KEY, value.jobId);
    } catch (reason) { if (token === generation.current) fail(reason); }
    finally { creating.current = false; }
  }

  async function convert() {
    if (!job?.scanReport) return;
    const token = generation.current;
    const seen = new Set<string>();
    setStage("converting");
    try {
      for (let index = 0; index < job.scanReport.batchCount; index += 1) {
        const source = assertFormulaBatch(await legacyEquationClient.readBatch(job.jobId, index), job.jobId, index);
        if (token !== generation.current) return;
        if (source.batchCount !== job.scanReport.batchCount) throw new Error("Formula batch count mismatch");
        const output: OmmlBatch = { protocolVersion: 1, jobId: job.jobId, batchIndex: index, batchCount: source.batchCount, formulas: [] };
        for (const formula of source.formulas) {
          if (seen.has(formula.formulaId)) throw new Error("Duplicate formula across batches");
          seen.add(formula.formulaId);
          const chosen = formula.riskLevel === "auto-replace"
            ? !deselected.has(formula.formulaId)
            : selected.has(formula.formulaId);
          if (!chosen || formula.riskLevel === "blocked" || !formula.mathMl) {
            output.formulas.push({ formulaId: formula.formulaId, status: "preserved", warnings: formula.warnings, errors: formula.errors });
            continue;
          }
          setStage("validating");
          try {
            const artifacts = mathMlToOmmlArtifacts(formula.mathMl, formula.displayMode);
            const validation = validateMathMlToOmml(formula.mathMl, artifacts.omml);
            if (!validation.valid) throw new Error(validation.errors.join("; "));
            output.formulas.push({ formulaId: formula.formulaId, status: "replaced", ommlBase64: artifacts.ommlBase64, warnings: [...formula.warnings, ...validation.warnings], errors: [] });
          } catch {
            output.formulas.push({ formulaId: formula.formulaId, status: "preserved", warnings: formula.warnings, errors: ["validation-failed"] });
          }
          setStage("converting");
        }
        await legacyEquationClient.submitBatch(job.jobId, index, output);
      }
      if (seen.size !== job.scanReport.detected) throw new Error("Formula count conservation failed");
      setStage("finalizing");
      const completed = await legacyEquationClient.finalize(job.jobId);
      if (token !== generation.current) return;
      setJob(completed); setStage("complete"); localStorage.removeItem(RECOVERY_KEY);
    } catch (reason) { if (token === generation.current) fail(reason); }
  }

  async function cancelJob() {
    if (!job) return;
    setStage("cancelling"); ++generation.current;
    try { setJob(await legacyEquationClient.cancel(job.jobId)); setStage("cancelled"); localStorage.removeItem(RECOVERY_KEY); }
    catch (reason) { fail(reason); }
  }

  const report = job?.scanReport;
  const labels: Array<[string, string, number]> = [
    [t("MathType 4 trở lên", "MathType 4 or later"), "Equation.DSMT4", report?.byProgId?.["Equation.DSMT4"] ?? 0],
    [t("Microsoft Equation Editor 3.0", "Microsoft Equation Editor 3.0"), "Equation.3", report?.byProgId?.["Equation.3"] ?? 0],
    ["Equation.2", "Equation.2", report?.byProgId?.["Equation.2"] ?? 0],
    ["Unknown OLE", "unknown", report?.byProgId?.unknown ?? 0],
  ];

  return <main className="legacy-equation-app" data-testid="legacy-equation-converter">
    <header><h1>{t("Chuyển công thức MathType cũ", "Convert legacy MathType equations")}</h1>
      <p role="note">{t("Equation Editor 3.0 chưa có corpus MTEF v3 thực để xác nhận.", "Equation Editor 3.0 has not been validated with a real MTEF v3 corpus.")}</p></header>
    <section aria-label={t("Chọn tài liệu Word", "Choose Word document")}>
      <button onClick={chooseInput}>{t("Chọn tài liệu Word", "Choose Word document")}</button><output>{inputPath}</output>
      <button onClick={chooseOutput} disabled={!inputPath}>{t("Chọn tệp đầu ra", "Choose output")}</button><output>{outputPath}</output>
      <button data-testid="scan" onClick={start} disabled={!inputPath || !outputPath || active}>{t("Quét công thức", "Scan equations")}</button>
    </section>
    <div role="progressbar" aria-valuemin={0} aria-valuemax={100} aria-valuenow={progress(stage)} aria-label={t("Tiến độ chuyển đổi", "Conversion progress")}>{progress(stage)}%</div>
    <p aria-live="polite">{({ scanning: t("Đang phân tích", "Analyzing"), converting: t("Đang chuyển đổi", "Converting"), validating: t("Đang xác thực", "Validating"), finalizing: t("Đang ghi tài liệu", "Writing document") } as Partial<Record<UiStage, string>>)[stage] ?? stage}</p>
    {error && <div ref={errorRef} tabIndex={-1} role="alert">{error}</div>}
    {report && <section><h2>{t("Kết quả quét", "Scan summary")}</h2>{labels.map(([label, key, count]) => <div key={key}>{label}: {count}</div>)}
      {(["auto-replace", "spot-check", "manual-review", "blocked"] as const).map((risk) => <div key={risk}>{risk}: {report.byRiskLevel?.[risk] ?? 0}</div>)}</section>}
    {batch && <section><h2>{t("Công thức", "Equations")} — {page + 1}/{batch.batchCount}</h2>
      <div className="legacy-formula-list">{batch.formulas.map((formula) => {
        const previewable = !!formula.mathMl && new TextEncoder().encode(formula.mathMl).byteLength <= MAX_PREVIEW_BYTES;
        return <article key={formula.formulaId}><label><input type="checkbox" disabled={formula.riskLevel === "blocked"}
          checked={formula.riskLevel === "auto-replace" ? !deselected.has(formula.formulaId) : selected.has(formula.formulaId)} onChange={(event) => {
            if (formula.riskLevel === "auto-replace") setDeselected((old) => { const next = new Set(old); event.target.checked ? next.delete(formula.formulaId) : next.add(formula.formulaId); return next; });
            else setSelected((old) => { const next = new Set(old); event.target.checked ? next.add(formula.formulaId) : next.delete(formula.formulaId); return next; });
          }}/>{formula.formulaId} · {formula.riskLevel}</label>
          <PreviewBoundary fallback={t("Không thể xem trước", "Preview unavailable")}><div className="legacy-math-preview">{previewable && formula.mathMl ? safeMathMlPreview(formula.mathMl) : t("Bản xem trước quá lớn", "Preview too large")}</div></PreviewBoundary></article>;
      })}</div>
      <nav><button disabled={page === 0} onClick={() => job && void loadBatch(job, page - 1)}>{t("Trước", "Previous")}</button><button disabled={page + 1 >= batch.batchCount} onClick={() => job && void loadBatch(job, page + 1)}>{t("Sau", "Next")}</button></nav>
      <button data-testid="convert" onClick={convert} disabled={stage !== "review"}>{t("Xác nhận chuyển đổi", "Confirm conversion")}</button></section>}
    {stage === "review" && report?.detected === 0 && <button data-testid="convert" onClick={convert}>{t("Xác nhận tạo báo cáo", "Confirm and create report")}</button>}
    {active && <button onClick={cancelJob}>{t("Hủy chuyển đổi", "Cancel conversion")}</button>}
    {(stage === "failed" || stage === "cancelled") && <button onClick={start}>{t("Thử lại", "Retry")}</button>}
    {stage === "complete" && job && <section><p>{job.outputPath}</p><button onClick={() => legacyEquationClient.openOutput(job.jobId, "word")}>{t("Mở tài liệu đầu ra", "Open output document")}</button><button onClick={() => legacyEquationClient.openOutput(job.jobId, "finder")}>Finder</button><button onClick={() => legacyEquationClient.openReport(job.jobId)}>{t("Mở báo cáo", "Open report")}</button></section>}
  </main>;
}

export default LegacyEquationConverterApp;
