import {
  Component,
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type ErrorInfo,
  type ReactNode,
} from "react";
import {
  AlertCircle,
  Check,
  FileCode2,
  FileText,
  Image as ImageIcon,
  LoaderCircle,
  Sigma,
  X,
} from "lucide-react";
import { MathPreview } from "../../components/MathPreview";
import { latexToSvg } from "../../export/latexToSvg";
import { latexLinesToOmmlArtifacts } from "../omml/latexToOmml";
import { onCurrentTauriWindowCloseRequested } from "../shared/tauriTransport";
import {
  createFormulaMetadata,
  type VisualTeXFormulaMetadata,
} from "../shared/formulaMetadata";
import { normalizeFormulaEditorDocument } from "../shared/formulaEditorDocument";
import { createUuid } from "../../runtime/browserCompatibility";
import { documentImportErrorMessage } from "./documentImportErrors";
import {
  cancelMacosDocumentImport,
  closeMacosDocumentImportWindow,
  commitMacosDocumentImport,
  getMacosDocumentImportRequest,
  type DocumentImportCommitItem,
  type MacosDocumentImportRequest,
} from "./documentImportClient";
import {
  mergeDocumentImportBlocks,
  parseLatexMarkdownDocument,
  type DocumentFormulaBlock,
  type DocumentFormulaOutputKind,
  type DocumentImportBlock,
  type DocumentImportSourceKind,
} from "./documentImportParser";

const MAX_WORD_REFERENCE_WIDTH_PT = 500;
const REFERENCE_FONT_SIZE_PT = 14;

class FormulaPreviewBoundary extends Component<
  { children: ReactNode; message: string },
  { failed: boolean }
> {
  state = { failed: false };

  static getDerivedStateFromError() {
    return { failed: true };
  }

  componentDidCatch(error: Error, info: ErrorInfo) {
    console.warn("VisualTeX document formula preview failed", error, info);
  }

  componentDidUpdate(previous: { children: ReactNode; message: string }) {
    if (this.state.failed && previous.children !== this.props.children) {
      this.setState({ failed: false });
    }
  }

  render() {
    if (this.state.failed) {
      return <span className="document-import-formula-error">{this.props.message}</span>;
    }
    return this.props.children;
  }
}

function clampFontSize(value: number, fallback = 12) {
  if (!Number.isFinite(value)) return fallback;
  return Math.min(512, Math.max(1, Math.round(value * 2) / 2));
}

function formulaCount(blocks: DocumentImportBlock[]) {
  return blocks.reduce((count, block) => count + (block.kind === "formula" ? 1 : 0), 0);
}

function textCharacterCount(blocks: DocumentImportBlock[]) {
  return blocks.reduce(
    (count, block) => count + (block.kind === "text" ? block.text.length : 0),
    0,
  );
}

function calculateReferenceGeometry(widthPx: number, heightPx: number, baselinePx: number) {
  const naturalWidthPt = widthPx * 0.75;
  const naturalHeightPt = heightPx * 0.75;
  const scale = Math.min(1, MAX_WORD_REFERENCE_WIDTH_PT / naturalWidthPt);
  const referenceWidthPt = naturalWidthPt * scale;
  const referenceHeightPt = naturalHeightPt * scale;
  const descentRatio = Math.max(0, Math.min(1, (heightPx - baselinePx) / heightPx));
  const referenceBaselinePt = -Math.max(
    0,
    Math.round(referenceHeightPt * descentRatio),
  );
  return { referenceWidthPt, referenceHeightPt, referenceBaselinePt };
}

async function prepareFormulaCommitItem(
  block: DocumentFormulaBlock,
  outputKind: DocumentFormulaOutputKind,
): Promise<DocumentImportCommitItem> {
  const formulaId = createUuid();
  const line = { id: createUuid(), latex: block.latex.trim() };
  if (!line.latex) throw new Error("存在空公式，请填写或删除后再插入。");
  const editorDocument = normalizeFormulaEditorDocument([line], "raw");
  const omml = latexLinesToOmmlArtifacts([line.latex], block.displayMode);

  const paragraphMetadata = {
    paragraphId: block.paragraphId,
    paragraphStyle: block.paragraphStyle,
    paragraphAlignment: block.paragraphAlignment,
    listKind: block.listKind,
    listLevel: block.listLevel,
    paragraphStart: block.paragraphStart,
    paragraphEnd: block.paragraphEnd,
  };

  let metadata: VisualTeXFormulaMetadata;
  if (outputKind === "image") {
    const svg = latexToSvg(line.latex, {
      displayMode: block.displayMode === "block",
      fontSizePt: REFERENCE_FONT_SIZE_PT,
      paddingPx: block.displayMode === "inline" ? 1 : 10,
      background: "transparent",
    });
    let pngBase64: string | undefined;
    try {
      const { svgToPng } = await import("../../export/svgToPng");
      pngBase64 = (
        await svgToPng(svg, { scale: 2, background: "transparent" })
      ).base64;
    } catch (reason) {
      console.warn(
        "VisualTeX document import could not create the optional PNG fallback",
        reason,
      );
    }
    const resolvedBaseline = svg.baseline ?? svg.height;
    const reference = calculateReferenceGeometry(
      svg.width,
      svg.height,
      resolvedBaseline,
    );
    metadata = createFormulaMetadata({
      formulaId,
      title: block.displayMode === "inline" ? "Imported inline formula" : "Imported display formula",
      lines: editorDocument.lines,
      codeFormat: editorDocument.codeFormat,
      displayMode: block.displayMode,
      numbered: block.displayMode === "block" && block.numbered,
      fontSizePt: block.fontSizePt,
      renderWidthPx: svg.width,
      renderHeightPx: svg.height,
      ...reference,
    });
    return {
      kind: "formula",
      formulaId,
      latex: line.latex,
      displayMode: block.displayMode,
      numbered: block.displayMode === "block" && block.numbered,
      fontSizePt: block.fontSizePt,
      metadata,
      ommlBase64: omml.ommlBase64,
      ommlDocxBase64: omml.ommlDocxBase64,
      svgBase64: svg.base64,
      ...(pngBase64 ? { pngBase64 } : {}),
      width: svg.width,
      height: svg.height,
      baseline: resolvedBaseline,
      ...paragraphMetadata,
    };
  }

  metadata = createFormulaMetadata({
    formulaId,
    title: block.displayMode === "inline" ? "Imported inline formula" : "Imported display formula",
    lines: editorDocument.lines,
    codeFormat: editorDocument.codeFormat,
    displayMode: block.displayMode,
    numbered: block.displayMode === "block" && block.numbered,
    fontSizePt: block.fontSizePt,
  });
  return {
    kind: "formula",
    formulaId,
    latex: line.latex,
    displayMode: block.displayMode,
    numbered: block.displayMode === "block" && block.numbered,
    fontSizePt: block.fontSizePt,
    metadata,
    ommlBase64: omml.ommlBase64,
    ommlDocxBase64: omml.ommlDocxBase64,
    ...paragraphMetadata,
  };
}

export function OfficeDocumentImportApp() {
  const sessionId = useMemo(
    () => new URLSearchParams(window.location.search).get("sessionId") ?? "",
    [],
  );
  const [request, setRequest] = useState<MacosDocumentImportRequest | null>(null);
  const [source, setSource] = useState("");
  const [sourceKind, setSourceKind] = useState<DocumentImportSourceKind>("auto");
  const [outputKind, setOutputKind] = useState<DocumentFormulaOutputKind>("omml");
  const [blocks, setBlocks] = useState<DocumentImportBlock[]>([]);
  const [loading, setLoading] = useState(true);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [toast, setToast] = useState("");
  const sourceRef = useRef<HTMLTextAreaElement>(null);
  const allowNativeCloseRef = useRef(false);
  const nativeCloseInFlightRef = useRef(false);

  useEffect(() => {
    if (!sessionId) {
      setError("缺少 Word 文档导入会话标识。");
      setLoading(false);
      return;
    }
    void getMacosDocumentImportRequest(sessionId)
      .then((value) => setRequest(value))
      .catch((reason) =>
        setError(documentImportErrorMessage(reason, "无法读取 Word 文档导入请求。")),
      )
      .finally(() => setLoading(false));
  }, [sessionId]);

  useEffect(() => {
    if (loading) return;
    const frame = window.requestAnimationFrame(() => sourceRef.current?.focus());
    return () => window.cancelAnimationFrame(frame);
  }, [loading]);

  const reparse = useCallback(
    (nextSource: string, nextKind = sourceKind) => {
      const parsed = parseLatexMarkdownDocument(
        nextSource,
        nextKind,
        request?.defaultFontSizePt ?? 12,
      );
      setBlocks((previous) => mergeDocumentImportBlocks(previous, parsed));
    },
    [request?.defaultFontSizePt, sourceKind],
  );

  const updateFormula = useCallback(
    (id: string, update: Partial<Omit<DocumentFormulaBlock, "id" | "kind">>) => {
      setBlocks((current) =>
        current.map((block) =>
          block.kind === "formula" && block.id === id ? { ...block, ...update } : block,
        ),
      );
    },
    [],
  );

  const handleSourceChange = (value: string) => {
    setSource(value);
    setError("");
    reparse(value);
  };

  const handleSourceKindChange = (value: DocumentImportSourceKind) => {
    setSourceKind(value);
    setError("");
    const parsed = parseLatexMarkdownDocument(
      source,
      value,
      request?.defaultFontSizePt ?? 12,
    );
    setBlocks((previous) => mergeDocumentImportBlocks(previous, parsed));
  };

  const cancel = async () => {
    if (busy || nativeCloseInFlightRef.current) return;
    nativeCloseInFlightRef.current = true;
    setBusy(true);
    setError("");
    try {
      if (sessionId) await cancelMacosDocumentImport(sessionId);
      allowNativeCloseRef.current = true;
      await closeMacosDocumentImportWindow();
    } catch (reason) {
      setError(documentImportErrorMessage(reason, "无法取消文档导入。"));
      setBusy(false);
      nativeCloseInFlightRef.current = false;
    }
  };

  const commit = async () => {
    if (!request || busy) return;
    if (!blocks.length || blocks.every((block) => block.kind === "text" && !block.text.trim())) {
      setError("请先粘贴包含文字或公式的 LaTeX/Markdown 内容。");
      return;
    }
    const formulas = blocks.filter(
      (block): block is DocumentFormulaBlock => block.kind === "formula",
    );
    if (formulas.some((block) => !block.latex.trim())) {
      setError("存在空公式，请填写公式内容后再插入。");
      return;
    }

    setBusy(true);
    setError("");
    setToast(outputKind === "omml" ? "正在生成 Word 原生公式…" : "正在生成 SVG 图片公式…");
    try {
      const preparedFormulas = await Promise.all(
        formulas.map((block) => prepareFormulaCommitItem(block, outputKind)),
      );
      let formulaIndex = 0;
      const items: DocumentImportCommitItem[] = blocks.map((block) => {
        if (block.kind === "text") {
          return {
            kind: "text",
            text: block.text,
            paragraphId: block.paragraphId,
            paragraphStyle: block.paragraphStyle,
            paragraphAlignment: block.paragraphAlignment,
            listKind: block.listKind,
            listLevel: block.listLevel,
            paragraphStart: block.paragraphStart,
            paragraphEnd: block.paragraphEnd,
          };
        }
        const prepared = preparedFormulas[formulaIndex];
        formulaIndex += 1;
        return prepared;
      });
      setToast("正在写入 Word…");
      await commitMacosDocumentImport(sessionId, { outputKind, items });
      allowNativeCloseRef.current = true;
      await closeMacosDocumentImportWindow();
    } catch (reason) {
      setError(documentImportErrorMessage(reason, "无法将内容插入 Word。"));
      setToast("");
      setBusy(false);
    }
  };

  useEffect(() => {
    if (!sessionId) return;
    let disposed = false;
    let unlisten: (() => void) | undefined;
    void onCurrentTauriWindowCloseRequested((event) => {
      if (disposed || allowNativeCloseRef.current) return;
      event.preventDefault();
      void cancel();
    })
      .then((dispose) => {
        if (disposed) dispose();
        else unlisten = dispose;
      })
      .catch((reason) => {
        setError(documentImportErrorMessage(reason, "无法注册文档导入窗口关闭处理。"));
      });
    return () => {
      disposed = true;
      unlisten?.();
    };
  }, [sessionId, busy]);

  if (loading) {
    return (
      <main className="document-import-state">
        <LoaderCircle className="is-spinning" />
        <span>正在准备 Word 文档导入器…</span>
      </main>
    );
  }

  if (!request) {
    return (
      <main className="document-import-state is-error" role="alert">
        <AlertCircle />
        <strong>无法打开文档导入器</strong>
        <p>{error || "Word 文档导入请求不存在或已经失效。"}</p>
      </main>
    );
  }

  const formulas = formulaCount(blocks);
  const textCharacters = textCharacterCount(blocks);

  return (
    <main className="document-import-app">
      <header className="document-import-header">
        <div>
          <span className="document-import-brand"><FileCode2 size={19} /> VisualTeX</span>
          <h1>插入 LaTeX / Markdown 文档</h1>
          <p>文字作为 Word 普通文本插入；每个公式保持独立，可单独编辑和调整字号。</p>
        </div>
        <button type="button" className="icon-button" onClick={() => void cancel()} disabled={busy} aria-label="关闭">
          <X size={19} />
        </button>
      </header>

      <section className="document-import-toolbar" aria-label="文档导入设置">
        <label>
          <span>输入类型</span>
          <select
            value={sourceKind}
            onChange={(event) => handleSourceKindChange(event.target.value as DocumentImportSourceKind)}
            disabled={busy}
          >
            <option value="auto">自动识别</option>
            <option value="markdown">Markdown</option>
            <option value="latex">LaTeX 文档</option>
          </select>
        </label>
        <fieldset className="document-import-output-kind">
          <legend>公式插入格式</legend>
          <label className={outputKind === "omml" ? "is-selected" : ""}>
            <input
              type="radio"
              name="document-formula-output"
              checked={outputKind === "omml"}
              onChange={() => setOutputKind("omml")}
              disabled={busy}
            />
            <Sigma size={17} />
            <span><strong>Word 原生 OMML</strong><small>可直接用 Word 字号和公式工具编辑</small></span>
          </label>
          <label className={outputKind === "image" ? "is-selected" : ""}>
            <input
              type="radio"
              name="document-formula-output"
              checked={outputKind === "image"}
              onChange={() => setOutputKind("image")}
              disabled={busy}
            />
            <ImageIcon size={17} />
            <span><strong>SVG 图片公式</strong><small>保持矢量清晰度，可双击回到 VisualTeX 编辑</small></span>
          </label>
        </fieldset>
        <div className="document-import-summary">
          <span><FileText size={15} /> {textCharacters} 个文字字符</span>
          <span><Sigma size={15} /> {formulas} 个独立公式</span>
        </div>
      </section>

      <section className="document-import-workspace">
        <div className="document-import-source-pane">
          <div className="document-import-pane-title">
            <strong>粘贴内容</strong>
            <span>支持行内与行间数学分隔符</span>
          </div>
          <textarea
            ref={sourceRef}
            value={source}
            onChange={(event) => handleSourceChange(event.target.value)}
            placeholder={
              "在这里粘贴 Markdown 或 LaTeX，例如：\n\n动量满足 $p=mv$。\n\n$$E=mc^2$$"
            }
            spellCheck={false}
            disabled={busy}
          />
        </div>

        <div className="document-import-preview-pane">
          <div className="document-import-pane-title">
            <strong>Word 插入预览</strong>
            <span>每张公式卡片都可以独立修改</span>
          </div>
          <div className="document-import-preview-document">
            {!blocks.length ? (
              <div className="document-import-empty">
                <FileText size={34} />
                <strong>尚无可预览内容</strong>
                <span>粘贴内容后，右侧会显示文本与公式的分段结果。</span>
              </div>
            ) : (
              blocks.map((block, index) =>
                block.kind === "text" ? (
                  <div
                    key={block.id}
                    className={`document-import-text-preview is-${block.paragraphStyle ?? "normal"} is-${block.listKind ?? "none"}`}
                    data-paragraph-start={block.paragraphStart ? "true" : undefined}
                  >
                    {block.paragraphStart && block.listKind === "bullet" ? (
                      <span className="document-import-list-marker">•</span>
                    ) : block.paragraphStart && block.listKind === "number" ? (
                      <span className="document-import-list-marker">1.</span>
                    ) : null}
                    {block.text}
                  </div>
                ) : (
                  <article
                    key={block.id}
                    className={`document-import-formula-card is-${block.displayMode}`}
                  >
                    <header>
                      <span>公式 {index + 1}</span>
                      <div>
                        <select
                          value={block.displayMode}
                          onChange={(event) =>
                            updateFormula(block.id, {
                              displayMode: event.target.value as "inline" | "block",
                              numbered:
                                event.target.value === "block" ? block.numbered : false,
                            })
                          }
                          disabled={busy}
                          aria-label="公式显示模式"
                        >
                          <option value="inline">行内公式</option>
                          <option value="block">行间公式</option>
                        </select>
                        {block.displayMode === "block" ? (
                          <label className="document-import-number-toggle">
                            <input
                              type="checkbox"
                              checked={block.numbered}
                              onChange={(event) =>
                                updateFormula(block.id, { numbered: event.target.checked })
                              }
                              disabled={busy}
                            />
                            <span>编号</span>
                          </label>
                        ) : null}
                        <label>
                          <span>字号</span>
                          <input
                            type="number"
                            min="1"
                            max="512"
                            step="0.5"
                            value={block.fontSizePt}
                            onChange={(event) =>
                              updateFormula(block.id, {
                                fontSizePt: clampFontSize(
                                  Number(event.target.value),
                                  block.fontSizePt,
                                ),
                              })
                            }
                            disabled={busy}
                          />
                          <span>pt</span>
                        </label>
                      </div>
                    </header>
                    <div className="document-import-formula-preview">
                      <FormulaPreviewBoundary message="公式暂时无法预览，请检查 LaTeX。">
                        <MathPreview latex={block.latex || "\\placeholder{}"} />
                      </FormulaPreviewBoundary>
                    </div>
                    <textarea
                      value={block.latex}
                      onChange={(event) => updateFormula(block.id, { latex: event.target.value })}
                      spellCheck={false}
                      disabled={busy}
                      aria-label="编辑公式 LaTeX"
                    />
                  </article>
                ),
              )
            )}
          </div>
        </div>
      </section>

      <footer className="document-import-footer">
        <div>
          {error ? (
            <span className="document-import-error" role="alert"><AlertCircle size={16} />{error}</span>
          ) : toast ? (
            <span className="document-import-progress"><LoaderCircle size={16} className="is-spinning" />{toast}</span>
          ) : (
            <span>插入后，每个公式都保留自己的字号和 VisualTeX 元数据。</span>
          )}
        </div>
        <button type="button" className="secondary-button" onClick={() => void cancel()} disabled={busy}>
          取消
        </button>
        <button
          type="button"
          className="primary-button"
          onClick={() => void commit()}
          disabled={busy || !blocks.length}
        >
          {busy ? <LoaderCircle size={16} className="is-spinning" /> : <Check size={16} />}
          插入 Word
        </button>
      </footer>
    </main>
  );
}
