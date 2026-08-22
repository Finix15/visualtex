import {
  Component,
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type ChangeEvent,
  type ErrorInfo,
  type ReactNode,
} from "react";
import {
  AlertCircle,
  Braces,
  Check,
  CheckCircle2,
  Eye,
  FileText,
  Image as ImageIcon,
  Upload,
  LoaderCircle,
  Sigma,
  X,
} from "lucide-react";
import { MathPreview } from "../../components/MathPreview";
import { onCurrentTauriWindowCloseRequested } from "../shared/tauriTransport";
import {
  createFormulaMetadata,
  type VisualTeXFormulaMetadata,
} from "../shared/formulaMetadata";
import {
  normalizeFormulaEditorDocument,
} from "../shared/formulaEditorDocument";
import {
  OFFICE_FORMULA_REFERENCE_FONT_SIZE_PT,
  renderOfficeFormulaArtifacts,
} from "../shared/formulaRenderArtifacts";
import { createUuid } from "../../runtime/browserCompatibility";
import { useEditorStore } from "../../stores/editorStore";
import { documentImportErrorMessage } from "./documentImportErrors";
import {
  cancelMacosDocumentImport,
  closeMacosDocumentImportWindow,
  commitMacosDocumentImport,
  focusMacosDocumentImportTarget,
  getMacosDocumentImportProgress,
  getMacosDocumentImportRequest,
  restoreMacosDocumentImportWindow,
  type DocumentImportCommitItem,
  type MacosDocumentImportRequest,
} from "./documentImportClient";
import {
  type ImportedDocumentFile,
  readDocumentImportFile,
} from "./documentImportFile.ts";
import {
  mergeDocumentImportBlocks,
  parseLatexMarkdownDocument,
  type DocumentFormulaBlock,
  type DocumentFormulaOutputKind,
  type DocumentImportBlock,
  type DocumentImportSourceKind,
} from "./documentImportParser";

const MAX_WORD_REFERENCE_WIDTH_PT = 500;
const WORD_IMAGE_VISUAL_SCALE = 1.1;
const REFERENCE_FONT_SIZE_PT = OFFICE_FORMULA_REFERENCE_FONT_SIZE_PT;

type ImportedFileState = Pick<
  ImportedDocumentFile,
  "name" | "encoding" | "size"
> & { modified: boolean };

function formatFileSize(bytes: number) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

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

function decodeUrlSafeBase64Utf8(value: string) {
  const normalized = value.replace(/-/g, "+").replace(/_/g, "/");
  const padded = normalized.padEnd(Math.ceil(normalized.length / 4) * 4, "=");
  const binary = atob(padded);
  const bytes = Uint8Array.from(binary, (character) => character.charCodeAt(0));
  return new TextDecoder().decode(bytes);
}

function ommlRetainsLiteralLatexCommand(ommlBase64: string, latex: string) {
  const commands = [...latex.matchAll(/\\([A-Za-z@]+)\b/g)]
    .map((match) => match[1])
    .filter((command, index, values) => values.indexOf(command) === index);
  if (!commands.length) return false;
  const omml = decodeUrlSafeBase64Utf8(ommlBase64);
  return commands.some((command) => omml.includes(`\\${command}`));
}

function calculateReferenceGeometry(widthPx: number, heightPx: number, baselinePx: number) {
  const naturalWidthPt = widthPx * 0.75 * WORD_IMAGE_VISUAL_SCALE;
  const naturalHeightPt = heightPx * 0.75 * WORD_IMAGE_VISUAL_SCALE;
  const scale = Math.min(1, MAX_WORD_REFERENCE_WIDTH_PT / naturalWidthPt);
  const referenceWidthPt = naturalWidthPt * scale;
  const referenceHeightPt = naturalHeightPt * scale;
  const descentRatio = Math.max(0, Math.min(1, (heightPx - baselinePx) / heightPx));
  // Keep the exported descent as a fractional 14 pt reference. Word accepts
  // only an integer Font.Position, so rounding here and again after scaling to
  // the requested font size makes short subscript formulas (for example L_z)
  // lose almost a full point relative to superscript formulas such as L^2.
  // Round exactly once at the final Word dispatch boundary instead.
  const referenceBaselinePt = -Math.max(
    0,
    referenceHeightPt * descentRatio,
  );
  return { referenceWidthPt, referenceHeightPt, referenceBaselinePt };
}

async function prepareFormulaArtifactCommitItem(
  block: DocumentFormulaBlock,
  outputKind: DocumentFormulaOutputKind,
): Promise<DocumentImportCommitItem> {
  const formulaId = createUuid();
  const line = { id: createUuid(), latex: block.latex.trim() };
  if (!line.latex) throw new Error("There is an empty formula, please fill it in or delete it and then insert it.");
  const editorDocument = normalizeFormulaEditorDocument([line], "raw");
  const { formulaLetterFont, formulaChineseFont } = useEditorStore.getState();
  const artifacts = renderOfficeFormulaArtifacts({
    lines: editorDocument.lines,
    codeFormat: editorDocument.codeFormat,
    displayMode: block.displayMode,
    host: "word",
    formulaLetterFont,
    formulaChineseFont,
  });
  const { canonicalLatex, svg } = artifacts;
  if (!artifacts.omml) {
    throw new Error("Unable to generate Word OMML formula artifact.");
  }
  const omml = artifacts.omml;
  if (ommlRetainsLiteralLatexCommand(omml.ommlBase64, canonicalLatex)) {
    throw new Error("The formula contains a custom command that is not recognized by the Word formula converter.");
  }

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
    const { svgToPng } = await import("../../export/svgToPng");
    const pngBase64 = (
      await svgToPng(svg, { scale: 2, background: "transparent" })
    ).base64;
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
      sourceLatex: canonicalLatex,
      displayMode: block.displayMode,
      numbered: block.displayMode === "block" && block.numbered,
      fontSizePt: block.fontSizePt,
      formulaLetterFont,
      formulaChineseFont,
      renderWidthPx: svg.width,
      renderHeightPx: svg.height,
      ...reference,
    });
    return {
      kind: "formula",
      formulaId,
      latex: canonicalLatex,
      displayMode: block.displayMode,
      numbered: block.displayMode === "block" && block.numbered,
      fontSizePt: block.fontSizePt,
      metadata,
      ommlBase64: omml.ommlBase64,
      ommlDocxBase64: omml.ommlDocxBase64,
      svgBase64: svg.base64,
      pngBase64,
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
    sourceLatex: canonicalLatex,
    displayMode: block.displayMode,
    numbered: block.displayMode === "block" && block.numbered,
    fontSizePt: block.fontSizePt,
    formulaLetterFont,
    formulaChineseFont,
  });
  return {
    kind: "formula",
    formulaId,
    latex: canonicalLatex,
    displayMode: block.displayMode,
    numbered: block.displayMode === "block" && block.numbered,
    fontSizePt: block.fontSizePt,
    metadata,
    ommlBase64: omml.ommlBase64,
    ommlDocxBase64: omml.ommlDocxBase64,
    ...paragraphMetadata,
  };
}

function formulaLiteralFallbackText(block: DocumentFormulaBlock) {
  const original = block.sourceText?.trim();
  if (original) return original;
  const latex = block.latex.trim();
  if (block.displayMode === "inline") return `\\(${latex}\\)`;
  if (/^\\begin\s*\{[^{}]+\}/.test(latex)) return latex;
  return `\\[\n${latex}\n\\]`;
}

async function prepareFormulaCommitItem(
  block: DocumentFormulaBlock,
  outputKind: DocumentFormulaOutputKind,
): Promise<DocumentImportCommitItem> {
  try {
    return await prepareFormulaArtifactCommitItem(block, outputKind);
  } catch (reason) {
    console.warn(
      "VisualTeX preserved an unsupported document formula as literal text",
      reason,
      block.latex,
    );
    const paragraphMetadata = block.paragraphId
      ? {
          paragraphId: block.paragraphId,
          paragraphStyle: block.paragraphStyle,
          paragraphAlignment: block.paragraphAlignment,
          listKind: block.listKind,
          listLevel: block.listLevel,
          paragraphStart: block.paragraphStart,
          paragraphEnd: block.paragraphEnd,
        }
      : {
          paragraphId: createUuid(),
          paragraphStyle: "code" as const,
          paragraphAlignment: "left" as const,
          listKind: "none" as const,
          listLevel: 0,
          paragraphStart: true,
          paragraphEnd: true,
        };
    return {
      kind: "text",
      text: formulaLiteralFallbackText(block),
      ...paragraphMetadata,
    };
  }
}

export function OfficeDocumentImportApp() {
  const language = useEditorStore((state) => state.language);
  const isEn = language === "en";
  const sessionId = useMemo(
    () => new URLSearchParams(window.location.search).get("sessionId") ?? "",
    [],
  );
  const [request, setRequest] = useState<MacosDocumentImportRequest | null>(null);
  const [source, setSource] = useState("");
  const [sourceKind, setSourceKind] = useState<DocumentImportSourceKind>("auto");
  const [outputKind, setOutputKind] = useState<DocumentFormulaOutputKind>("omml");
  const [blocks, setBlocks] = useState<DocumentImportBlock[]>([]);
  const [importedFile, setImportedFile] = useState<ImportedFileState | null>(null);
  const [loading, setLoading] = useState(true);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [toast, setToast] = useState("");
  const fileInputRef = useRef<HTMLInputElement>(null);
  const sourceRef = useRef<HTMLTextAreaElement>(null);
  const allowNativeCloseRef = useRef(false);
  const nativeCloseInFlightRef = useRef(false);

  useEffect(() => {
    if (!sessionId) {
      setError((isEn ? "The Word document import session ID is missing." : "ID phiên nhập tài liệu Word bị thiếu."));
      setLoading(false);
      return;
    }
    void getMacosDocumentImportRequest(sessionId)
      .then((value) => setRequest(value))
      .catch((reason) =>
        setError(documentImportErrorMessage(reason, (isEn ? "The Word document import request cannot be read." : "Yêu cầu nhập tài liệu Word không thể đọc được."))),
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
        current.map((block) => {
          if (block.kind !== "formula" || block.id !== id) return block;
          const updated = { ...block, ...update };
          if (
            ("latex" in update && update.latex !== block.latex) ||
            ("displayMode" in update && update.displayMode !== block.displayMode)
          ) {
            delete updated.sourceText;
          }
          return updated;
        }),
      );
    },
    [],
  );

  const handleSourceChange = (value: string) => {
    setSource(value);
    setImportedFile((current) =>
      current ? { ...current, modified: true } : current,
    );
    setError("");
    reparse(value);
  };

  const handleDocumentFile = async (file: File | null) => {
    if (!file || busy) return;
    setError("");
    setToast((isEn ? "Reading document source code..." : "Đọc mã nguồn tài liệu..."));
    try {
      const imported = await readDocumentImportFile(file);
      setSource(imported.source);
      setSourceKind(imported.format);
      setImportedFile({
        name: imported.name,
        encoding: imported.encoding,
        size: imported.size,
        modified: false,
      });
      const parsed = parseLatexMarkdownDocument(
        imported.source,
        imported.format,
        request?.defaultFontSizePt ?? 12,
      );
      setBlocks((previous) => mergeDocumentImportBlocks(previous, parsed));
      setToast(
        (isEn ? `loaded${imported.name} · ${imported.encoding} · ${formatFileSize(imported.size)}` : `đã tải${imported.name} · ${imported.encoding} · ${formatFileSize(imported.size)}`),
      );
      window.requestAnimationFrame(() => sourceRef.current?.focus());
    } catch (reason) {
      setError(documentImportErrorMessage(reason, (isEn ? "Unable to read the document source file." : "Không thể đọc được file nguồn tài liệu.")));
      setToast("");
    } finally {
      if (fileInputRef.current) fileInputRef.current.value = "";
    }
  };

  const handleFileInput = (event: ChangeEvent<HTMLInputElement>) => {
    void handleDocumentFile(event.currentTarget.files?.[0] ?? null);
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
      setError(documentImportErrorMessage(reason, (isEn ? "Unable to cancel document import." : "Không thể hủy nhập tài liệu.")));
      setBusy(false);
      nativeCloseInFlightRef.current = false;
    }
  };

  const commit = async () => {
    if (!request || busy) return;
    if (!blocks.length || blocks.every((block) => block.kind === "text" && !block.text.trim())) {
      setError((isEn ? "Please paste the LaTeX/Markdown content containing text or formulas first." : "Vui lòng dán nội dung LaTeX/Markdown có chứa văn bản hoặc công thức trước."));
      return;
    }
    const formulas = blocks.filter(
      (block): block is DocumentFormulaBlock => block.kind === "formula",
    );
    if (formulas.some((block) => !block.latex.trim())) {
      setError((isEn ? "There is an empty formula. Please fill in the formula content before inserting it." : "Có công thức trống. Hãy điền nội dung công thức trước khi chèn vào."));
      return;
    }

    setBusy(true);
    setError("");
    setToast(outputKind === "omml" ? (isEn ? "Generating Word native formulas..." : "Tạo công thức gốc trong Word...") : (isEn ? "Generating SVG image formula..." : "Đang tạo công thức ảnh SVG..."));
    let importerHidden = false;
    let progressTimer: number | undefined;
    let progressRequestInFlight = false;
    try {
      await focusMacosDocumentImportTarget();
      importerHidden = true;
      const preparedFormulas = await Promise.all(
        formulas.map(async (block, index) => {
          try {
            return await prepareFormulaCommitItem(block, outputKind);
          } catch (reason) {
            const detail = documentImportErrorMessage(
              reason,
              (isEn ? "Unknown formula conversion error." : "Lỗi chuyển đổi công thức không xác định."),
            );
            const preview = block.latex.trim().replace(/\s+/g, " ").slice(0, 120);
            throw new Error(
              (isEn ? `formula${index + 1}Generation failed:${detail}${preview ? `（${preview}）` : ""}` : `công thức${index + 1}Thế hệ không thành công:${detail}${preview ? `（${preview}）` : ""}`),
              { cause: reason },
            );
          }
        }),
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
      const literalFallbackCount = preparedFormulas.filter(
        (item) => item.kind === "text",
      ).length;
      setToast(
        literalFallbackCount > 0
          ? (isEn ? `is writing to Word(${literalFallbackCount}unsupported fragments are retained as original text)…` : `đang viết vào Word(${literalFallbackCount}những đoạn không được hỗ trợ sẽ được giữ lại dưới dạng văn bản gốc)…`)
          : (isEn ? `Writing Word: 0/${items.length}` : `Viết chữ: 0/${items.length}`),
      );
      progressTimer = window.setInterval(() => {
        if (progressRequestInFlight) return;
        progressRequestInFlight = true;
        void getMacosDocumentImportProgress(sessionId)
          .then((progress) => {
            if (progress.total > 0 && progress.stage === "inserting") {
              setToast((isEn ? `is writing to Word:${progress.current}/${progress.total}` : `đang viết vào Word:${progress.current}/${progress.total}`));
            }
          })
          .catch(() => undefined)
          .finally(() => {
            progressRequestInFlight = false;
          });
      }, 120);
      await commitMacosDocumentImport(sessionId, { outputKind, items });
      if (progressTimer !== undefined) window.clearInterval(progressTimer);
      setToast((isEn ? `Completed:${items.length}/${items.length}` : `Đã hoàn thành:${items.length}/${items.length}`));
      allowNativeCloseRef.current = true;
      await closeMacosDocumentImportWindow();
    } catch (reason) {
      if (progressTimer !== undefined) window.clearInterval(progressTimer);
      if (importerHidden) {
        await restoreMacosDocumentImportWindow().catch(() => undefined);
      }
      setError(documentImportErrorMessage(reason, (isEn ? "Unable to insert content into Word." : "Không chèn được nội dung vào Word.")));
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
        setError(documentImportErrorMessage(reason, (isEn ? "Unable to register document import window closing processing." : "Không thể đăng ký xử lý đóng cửa sổ nhập tài liệu.")));
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
        <span>{isEn ? "Preparing Word document importer…" : "Đang chuẩn bị nhập tài liệu Word…"}</span>
      </main>
    );
  }

  if (!request) {
    return (
      <main className="document-import-state is-error" role="alert">
        <AlertCircle />
        <strong>{isEn ? "Unable to open document importer" : "Không thể mở trình nhập tài liệu"}</strong>
        <p>{error || (isEn ? "The Word document import request does not exist or has expired." : "Yêu cầu nhập tài liệu Word không tồn tại hoặc đã hết hạn.")}</p>
      </main>
    );
  }

  const formulas = formulaCount(blocks);
  const textCharacters = textCharacterCount(blocks);

  return (
    <main className="doc-import-shell macos-doc-import">
      <header className="doc-import-toolbar">
        <div className="doc-import-title-block">
          <FileText size={20} />
          <div>
            <strong>{isEn ? "Word document batch import" : "Nhập hàng loạt tài liệu Word"}</strong>
            <span>{isEn ? "Edit the source code on the left, view and adjust the final Word structure in real time on the right" : "Chỉnh sửa mã nguồn bên trái, xem và điều chỉnh cấu trúc Word cuối cùng theo thời gian thực bên phải"}</span>
          </div>
        </div>
        <div className="doc-import-options">
          <label>
            <span>{isEn ? "Source format" : "Định dạng nguồn"}</span>
            <select
              value={sourceKind}
              onChange={(event) =>
                handleSourceKindChange(event.target.value as DocumentImportSourceKind)
              }
              disabled={busy}
            >
              <option value="auto">{isEn ? "Automatic identification" : "Nhận dạng tự động"}</option>
              <option value="latex">LaTeX</option>
              <option value="markdown">Markdown</option>
            </select>
          </label>
          <label>
            <span>{isEn ? "Formula format" : "Định dạng công thức"}</span>
            <select
              value={outputKind}
              onChange={(event) =>
                setOutputKind(event.target.value as DocumentFormulaOutputKind)
              }
              disabled={busy}
            >
              <option value="omml">{isEn ? "Word native OMML" : "Từ OMML gốc"}</option>
              <option value="image">{isEn ? "SVG image formula" : "Công thức ảnh SVG"}</option>
            </select>
          </label>
          <button
            type="button"
            className="doc-import-secondary doc-import-file-button"
            onClick={() => fileInputRef.current?.click()}
            disabled={busy}
            title={isEn ? "Import a single LaTeX or Markdown file" : "Nhập một tệp LaTeX hoặc Markdown"}
          >
            <Upload size={16} />
            {isEn ? "Import .tex / .md" : "Nhập .tex / .md"}</button>
          <input
            ref={fileInputRef}
            type="file"
            accept=".tex,.md,.markdown,text/x-tex,text/markdown"
            aria-label={isEn ? "Import LaTeX or Markdown files" : "Nhập tệp LaTeX hoặc Markdown"}
            hidden
            onChange={handleFileInput}
          />
        </div>
      </header>

      <section className="doc-import-workspace">
        <article className="doc-import-pane source-pane">
          <div className="doc-import-pane-header">
            <div className="doc-import-pane-heading">
              <span className="doc-import-pane-icon" aria-hidden="true">
                <Braces size={16} />
              </span>
              <div>
                <strong>{isEn ? "LaTeX/Markdown source code" : "Mã nguồn LaTeX/Markdown"}</strong>
                <small>{isEn ? "supports text, titles, lists, theorems, quotes, code blocks and mixed formulas" : "hỗ trợ văn bản, tiêu đề, danh sách, định lý, dấu ngoặc kép, khối mã và công thức hỗn hợp"}</small>
              </div>
            </div>
            <div className="doc-import-source-meta">
              {importedFile ? (
                <span
                  className="doc-import-file-chip"
                  title={`${importedFile.name} · ${importedFile.encoding} · ${formatFileSize(importedFile.size)}`}
                >
                  <FileText size={12} />
                  <span>{importedFile.name}</span>
                  <small>
                    {importedFile.encoding}
                    {importedFile.modified ? (isEn ? "· Edited" : "· Đã chỉnh sửa") : ""}
                  </small>
                </span>
              ) : null}
              <span className="doc-import-pane-stat">
                {source.length.toLocaleString()} {isEn ? "character" : "ký tự"}</span>
            </div>
          </div>
          <textarea
            ref={sourceRef}
            value={source}
            onChange={(event) => handleSourceChange(event.target.value)}
            placeholder={isEn
              ? "Paste LaTeX or Markdown here, for example: Inline formula $E=mc^2$."
              : "Dán LaTeX hoặc Markdown vào đây, ví dụ: công thức cùng dòng $E=mc^2$."}
            spellCheck={false}
            autoCapitalize="off"
            autoCorrect="off"
            disabled={busy}
            aria-label={isEn ? "Document source code" : "Mã nguồn tài liệu"}
          />
        </article>

        <article className="doc-import-pane preview-pane">
          <div className="doc-import-pane-header">
            <div className="doc-import-pane-heading">
              <span className="doc-import-pane-icon is-preview" aria-hidden="true">
                <Eye size={16} />
              </span>
              <div>
                <strong>{isEn ? "Word structure preview" : "Xem trước cấu trúc từ"}</strong>
                <small>{isEn ? "Formula cards can individually adjust within/between lines, numbering and font size" : "Thẻ công thức có thể điều chỉnh riêng lẻ trong/giữa các dòng, đánh số và cỡ chữ"}</small>
              </div>
            </div>
            <div className="doc-import-preview-counts" aria-label={isEn ? "Preview statistics" : "Thống kê xem trước"}>
              <span>{blocks.length} {isEn ? "block" : ""}</span>
              <span>{textCharacters} {isEn ? "words" : "từ"}</span>
              <span>{formulas} {isEn ? "formula" : "công thức"}</span>
            </div>
          </div>
          <div className="doc-import-preview-scroll">
            <div className="doc-import-preview-document">
              {!blocks.length ? (
                <div className="document-import-empty">
                  <FileText size={34} />
                  <strong>{isEn ? "Waiting for document content" : "Đang chờ nội dung tài liệu"}</strong>
                  <span>{isEn ? "After pasting the content on the left, a Word structure preview will be generated in real time." : "Sau khi dán nội dung bên trái, bản xem trước cấu trúc Word sẽ được tạo theo thời gian thực."}</span>
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
                        <span>{isEn ? "formula" : "công thức"}{index + 1}</span>
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
                            aria-label={isEn ? "Formula display mode" : "Chế độ hiển thị công thức"}
                          >
                            <option value="inline">{isEn ? "Inline formula" : "Công thức nội tuyến"}</option>
                            <option value="block">{isEn ? "Interline formula" : "Công thức xen kẽ"}</option>
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
                              <span>{isEn ? "No." : "Không."}</span>
                            </label>
                          ) : null}
                          <label>
                            <span>{isEn ? "Font size" : "Cỡ chữ"}</span>
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
                        <FormulaPreviewBoundary message={isEn ? "The formula cannot be previewed temporarily, please check LaTeX." : "Tạm thời không thể xem trước công thức, vui lòng kiểm tra LaTeX."}>
                          <MathPreview latex={block.latex || "\\placeholder{}"} />
                        </FormulaPreviewBoundary>
                      </div>
                      <textarea
                        value={block.latex}
                        onChange={(event) => updateFormula(block.id, { latex: event.target.value })}
                        spellCheck={false}
                        disabled={busy}
                        aria-label={isEn ? "Edit formula LaTeX" : "Chỉnh sửa công thức LaTeX"}
                      />
                    </article>
                  ),
                )
              )}
            </div>
          </div>
        </article>
      </section>

      <footer className="doc-import-footer">
        <div className="doc-import-messages">
          {error ? (
            <span className="error" role="alert"><AlertCircle size={15} />{error}</span>
          ) : toast ? (
            <span><LoaderCircle size={15} className="is-spinning" />{toast}</span>
          ) : (
            <span className="ok">
              <CheckCircle2 size={15} />
              {isEn ? "The preview parsing is normal; after clicking Import, it will switch back to Word and display the insertion progress in real time." : "Việc phân tích bản xem trước là bình thường; sau khi nhấn Import sẽ chuyển về Word và hiển thị tiến trình chèn theo thời gian thực."}</span>
          )}
        </div>
        <div className="doc-import-actions">
          <button type="button" className="doc-import-secondary" onClick={() => void cancel()} disabled={busy}>
            <X size={16} />{isEn ? "Cancel" : "Hủy bỏ"}</button>
          <button
            type="button"
            className="doc-import-primary"
            onClick={() => void commit()}
            disabled={busy || !blocks.length}
          >
            {busy ? <LoaderCircle size={16} className="is-spinning" /> : <Check size={16} />}
            {busy ? (isEn ? "Importing..." : "Đang nhập...") : (isEn ? "Import into Word" : "Nhập vào Word")}
          </button>
        </div>
      </footer>
    </main>
  );
}
