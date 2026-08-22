import { useEffect, useMemo, useRef, useState } from "react";
import { createUuid } from "../../runtime/browserCompatibility";
import {
  Braces,
  CheckCircle2,
  Eye,
  FileText,
  FolderOpen,
  LoaderCircle,
  TriangleAlert,
  X,
} from "lucide-react";
import { readErrorMessage } from "../../errors/readErrorMessage";
import { latexToSvg } from "../../export/latexToSvg";
import {
  applyDocumentTheme,
  normalizeSynchronizedTheme,
  readSynchronizedTheme,
  subscribeSynchronizedTheme,
} from "../../themeSync";
import {
  closeOfficeSessionWindow,
  getOfficeSession,
  getOfficeTheme,
  saveOfficeSessionKeepalive,
  updateOfficeSession,
  type OfficeFormulaSession,
} from "../api/sessionClient";
import {
  readDocumentImportFile,
  type ImportedDocumentFile,
} from "./documentImportFile";
import {
  parseDocumentImport,
  type DocumentImportBlock,
  type DocumentImportRun,
  type DocumentObjectMode,
  type DocumentSourceFormat,
  type ParsedDocumentImport,
} from "./documentImportParser";
import "./documentImport.css";
import { useEditorStore } from "../../stores/editorStore";

function sessionIdFromLocation() {
  const match = window.location.pathname.match(/\/dialog\/([0-9a-f-]{36})/i);
  if (!match) throw new Error("The bulk import window is missing a valid Session id.");
  return match[1].toLowerCase();
}

function formatFromSession(session: OfficeFormulaSession): DocumentSourceFormat {
  if (session.codeFormat === "markdown-document") return "markdown";
  if (session.codeFormat === "latex-document") return "latex";
  return "auto";
}

function FormulaPreview({ latex, display }: { latex: string; display: boolean }) {
  const rendered = useMemo(() => {
    try {
      return {
        svg: latexToSvg(latex, {
          displayMode: display,
          fontSizePt: 13,
          paddingPx: display ? 4 : 1,
          background: "transparent",
        }).svg,
        error: "",
      };
    } catch (error) {
      return {
        svg: "",
        error: readErrorMessage(error, "Formula preview failed."),
      };
    }
  }, [display, latex]);

  if (rendered.error) {
    return (
      <span className={display ? "doc-import-formula-error display" : "doc-import-formula-error"}>
        <code>{latex}</code>
        <small>{rendered.error}</small>
      </span>
    );
  }
  return (
    <span
      className={display ? "doc-import-formula display" : "doc-import-formula"}
      dangerouslySetInnerHTML={{ __html: rendered.svg }}
    />
  );
}

function InlineRuns({ runs }: { runs: DocumentImportRun[] }) {
  return (
    <>
      {runs.map((run, index) => {
        if (run.kind === "formula") {
          return <FormulaPreview key={index} latex={run.latex} display={run.display} />;
        }
        const className = [
          run.bold ? "bold" : "",
          run.italic ? "italic" : "",
          run.code ? "code" : "",
          run.strike ? "strike" : "",
          run.underline ? "underline" : "",
        ]
          .filter(Boolean)
          .join(" ");
        return (
          <span key={index} className={className || undefined}>
            {run.text}
          </span>
        );
      })}
    </>
  );
}

function PreviewBlock({ block }: { block: DocumentImportBlock }) {
  if (block.kind === "display") {
    const formula = block.runs.find((run) => run.kind === "formula");
    return formula?.kind === "formula" ? (
      <div className="doc-import-display-row">
        <FormulaPreview latex={formula.latex} display />
      </div>
    ) : null;
  }
  if (block.kind === "code") {
    const text = block.runs.map((run) => (run.kind === "text" ? run.text : run.latex)).join("");
    return <pre className="doc-import-code-block">{text}</pre>;
  }
  if (block.kind === "heading") {
    const content = <InlineRuns runs={block.runs} />;
    switch (Math.min(6, Math.max(1, block.level))) {
      case 1:
        return <h1 className="doc-import-heading">{content}</h1>;
      case 2:
        return <h2 className="doc-import-heading">{content}</h2>;
      case 3:
        return <h3 className="doc-import-heading">{content}</h3>;
      case 4:
        return <h4 className="doc-import-heading">{content}</h4>;
      case 5:
        return <h5 className="doc-import-heading">{content}</h5>;
      default:
        return <h6 className="doc-import-heading">{content}</h6>;
    }
  }
  if (block.kind === "quote") {
    return (
      <blockquote className="doc-import-quote">
        <InlineRuns runs={block.runs} />
      </blockquote>
    );
  }
  if (block.kind === "bullet" || block.kind === "numbered") {
    return (
      <div
        className={`doc-import-list-row ${block.kind}`}
        style={{ paddingInlineStart: `${18 + block.level * 24}px` }}
      >
        <span className="doc-import-list-marker">
          {block.kind === "bullet" ? "•" : `${block.level + 1}.`}
        </span>
        <span><InlineRuns runs={block.runs} /></span>
      </div>
    );
  }
  return (
    <p className="doc-import-paragraph">
      <InlineRuns runs={block.runs} />
    </p>
  );
}

function PreviewPane({ parsed }: { parsed: ParsedDocumentImport }) {
  return (
    <div className="doc-import-preview-stage">
      <div className="doc-import-preview-caption" aria-hidden="true">
        <span>
          <Eye size={14} />
          "Word page preview"</span>
        <span>A4 · Real-time structure</span>
      </div>
      <div className="doc-import-preview-document" role="document">
        <div className="doc-import-paper-content">
          {parsed.blocks.length > 0 ? (
            parsed.blocks.map((block) => (
              <PreviewBlock key={block.id} block={block} />
            ))
          ) : (
            <div className="doc-import-paper-empty">
              <FileText size={28} />
              <strong>Waiting for document content</strong>
              <span>After entering or pasting LaTeX or Markdown on the left, a preview of the Word structure will be generated in real time.</span>
            </div>
          )}
        </div>
        <div className="doc-import-page-footer" aria-hidden="true">
          <span>VisualTeX document preview</span>
          <span>1</span>
        </div>
      </div>
    </div>
  );
}

type ImportedFileState = Pick<ImportedDocumentFile, "name" | "encoding" | "size"> & {
  modified: boolean;
};

function formatFileSize(size: number) {
  if (size < 1024) return `${size} B`;
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
  return `${(size / (1024 * 1024)).toFixed(1)} MB`;
}

export function DocumentImportApp() {
  const language = useEditorStore((state) => state.language);
  const isEn = language === "en";
  const sessionId = useMemo(sessionIdFromLocation, []);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const finalizedRef = useRef(false);
  const [session, setSession] = useState<OfficeFormulaSession | null>(null);
  const [source, setSource] = useState("");
  const [format, setFormat] = useState<DocumentSourceFormat>("auto");
  const [objectMode, setObjectMode] = useState<DocumentObjectMode>("wordOmml");
  const [loading, setLoading] = useState(true);
  const [busy, setBusy] = useState(false);
  const [fileBusy, setFileBusy] = useState(false);
  const [importedFile, setImportedFile] = useState<ImportedFileState | null>(null);
  const [loadError, setLoadError] = useState("");

  useEffect(() => {
    let disposed = false;
    const applyTheme = (value: unknown) => applyDocumentTheme(normalizeSynchronizedTheme(value));
    applyTheme(readSynchronizedTheme());
    const unsubscribe = subscribeSynchronizedTheme(applyTheme);
    const sync = async () => {
      try {
        const status = await getOfficeTheme();
        if (!disposed) applyTheme(status.theme);
      } catch {
        // Keep the last synchronized theme while the companion restarts.
      }
    };
    void sync();
    const interval = window.setInterval(() => void sync(), 500);
    return () => {
      disposed = true;
      unsubscribe();
      window.clearInterval(interval);
    };
  }, []);

  useEffect(() => {
    let disposed = false;
    void getOfficeSession(sessionId)
      .then((next) => {
        if (disposed) return;
        setSession(next);
        setSource(next.lines[0]?.latex ?? "");
        setFormat(formatFromSession(next));
        setObjectMode(next.objectMode === "nativeOle" ? "nativeOle" : "wordOmml");
        setLoading(false);
      })
      .catch((error) => {
        if (disposed) return;
        setLoadError(readErrorMessage(error, (isEn ? "Unable to read bulk import Session." : "Không thể đọc Phiên nhập hàng loạt.")));
        setLoading(false);
      });
    return () => {
      disposed = true;
    };
  }, [sessionId]);

  const preview = useMemo(() => {
    try {
      return { parsed: parseDocumentImport(source, format), error: "" };
    } catch (error) {
      return { parsed: null, error: readErrorMessage(error, (isEn ? "The current document cannot be parsed." : "Không thể phân tích cú pháp tài liệu hiện tại.")) };
    }
  }, [format, source]);

  useEffect(() => {
    const cancelOnClose = () => {
      if (finalizedRef.current || !session) return;
      void saveOfficeSessionKeepalive(sessionId, {
        status: "cancelled",
        explicitCancel: true,
        error: null,
      });
    };
    window.addEventListener("beforeunload", cancelOnClose);
    return () => window.removeEventListener("beforeunload", cancelOnClose);
  }, [session, sessionId]);

  const cancel = async () => {
    if (busy) return;
    finalizedRef.current = true;
    try {
      await updateOfficeSession(sessionId, {
        status: "cancelled",
        explicitCancel: true,
        error: null,
      });
    } finally {
      await closeOfficeSessionWindow(sessionId).catch(() => undefined);
    }
  };

  const commit = async () => {
    if (!session || !preview.parsed || preview.error || busy) return;
    setBusy(true);
    try {
      const lineId = session.lines[0]?.id || createUuid();
      const serializedDocument = JSON.stringify(preview.parsed);
      if (serializedDocument.length > 5_000_000) {
        throw new Error((isEn ? "The parsed document structure exceeds 5 MB and cannot be submitted to Word. Please split and import." : "Cấu trúc tài liệu được phân tích cú pháp vượt quá 5 MB và không thể gửi tới Word. Vui lòng chia nhỏ và nhập khẩu."));
      }
      await updateOfficeSession(sessionId, {
        title: (isEn ? "Word document batch import" : "Nhập hàng loạt tài liệu Word"),
        lines: [{ id: lineId, latex: serializedDocument }],
        activeLineId: lineId,
        codeFormat: "visualtex-document-json",
        objectMode,
        displayMode: "block",
        numbered: false,
        dirty: true,
        status: "committing",
        explicitCancel: false,
        error: null,
      });
      finalizedRef.current = true;
      await closeOfficeSessionWindow(sessionId).catch(() => undefined);
    } catch (error) {
      setLoadError(readErrorMessage(error, (isEn ? "Unable to submit document to Word." : "Không thể gửi tài liệu tới Word.")));
      setBusy(false);
    }
  };

  const openFile = async (file: File) => {
    if (fileBusy || busy) return;
    setFileBusy(true);
    setLoadError("");
    try {
      const imported = await readDocumentImportFile(file);
      setSource(imported.source);
      setFormat(imported.format);
      setImportedFile({
        name: imported.name,
        encoding: imported.encoding,
        size: imported.size,
        modified: false,
      });
    } catch (error) {
      setLoadError(readErrorMessage(error, (isEn ? "The selected file cannot be read." : "Không thể đọc được tập tin đã chọn.")));
    } finally {
      setFileBusy(false);
    }
  };

  if (loading) {
    return (
      <main className="doc-import-loading">
        <LoaderCircle className="spin" />
        <span>{isEn ? "Opening Word document importer…" : "Đang mở trình nhập tài liệu Word…"}</span>
      </main>
    );
  }
  if (!session) {
    return (
      <main className="doc-import-loading error">
        <TriangleAlert />
        <span>{loadError || (isEn ? "Unable to read bulk import Session." : "Không thể đọc Phiên nhập hàng loạt.")}</span>
      </main>
    );
  }

  return (
    <main className="doc-import-shell">
      <header className="doc-import-toolbar">
        <div className="doc-import-title-block">
          <FileText size={20} />
          <div>
            <strong>{isEn ? "Word document batch import" : "Nhập hàng loạt tài liệu Word"}</strong>
            <span>{isEn ? "Edit the source code on the left and view the Word import structure in real time on the right" : "Chỉnh sửa mã nguồn ở bên trái và xem cấu trúc nhập Word theo thời gian thực ở bên phải"}</span>
          </div>
        </div>
        <div className="doc-import-options">
          <label>
            <span>{isEn ? "Source format" : "Định dạng nguồn"}</span>
            <select value={format} onChange={(event) => setFormat(event.target.value as DocumentSourceFormat)}>
              <option value="auto">{isEn ? "Automatic identification" : "Nhận dạng tự động"}</option>
              <option value="latex">LaTeX</option>
              <option value="markdown">Markdown</option>
            </select>
          </label>
          <label>
            <span>{isEn ? "Formula format" : "Định dạng công thức"}</span>
            <select value={objectMode} onChange={(event) => setObjectMode(event.target.value as DocumentObjectMode)}>
              <option value="wordOmml">{isEn ? "Word native OMML" : "Từ OMML gốc"}</option>
              <option value="nativeOle">VisualTeX OLE</option>
            </select>
          </label>
          <button
            className="doc-import-secondary doc-import-file-button"
            onClick={() => fileInputRef.current?.click()}
            disabled={busy || fileBusy}
            title={isEn ? "Import a single LaTeX or Markdown file" : "Nhập một tệp LaTeX hoặc Markdown"}
          >
            {fileBusy ? <LoaderCircle size={16} className="spin" /> : <FolderOpen size={16} />}
            {fileBusy ? (isEn ? "Reading..." : "Đang đọc...") : (isEn ? "Import .tex / .md" : "Nhập .tex / .md")}
          </button>
          <input
            ref={fileInputRef}
            type="file"
            accept=".tex,.md,.markdown,text/x-tex,text/markdown"
            aria-label={isEn ? "Import LaTeX or Markdown files" : "Nhập tệp LaTeX hoặc Markdown"}
            hidden
            onChange={(event) => {
              const file = event.target.files?.[0];
              if (file) void openFile(file);
              event.currentTarget.value = "";
            }}
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
                <small>{isEn ? "supports body text, titles, lists, quotes, code blocks and mixed formulas" : "hỗ trợ nội dung văn bản, tiêu đề, danh sách, dấu ngoặc kép, khối mã và công thức hỗn hợp"}</small>
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
                  <small>{importedFile.encoding}{importedFile.modified ? (isEn ? "· Edited" : "· Đã chỉnh sửa") : ""}</small>
                </span>
              ) : null}
              <span className="doc-import-pane-stat">
                {source.length.toLocaleString()} {isEn ? "character" : "ký tự"}</span>
            </div>
          </div>
          <textarea
            value={source}
            placeholder={isEn
              ? "Paste LaTeX or Markdown here, for example: Inline formula $E=mc^2$."
              : "Dán LaTeX hoặc Markdown vào đây, ví dụ: công thức cùng dòng $E=mc^2$."}
            onChange={(event) => {
              setSource(event.target.value);
              setImportedFile((current) =>
                current && !current.modified ? { ...current, modified: true } : current,
              );
            }}
            spellCheck={false}
            autoCapitalize="off"
            autoCorrect="off"
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
                <small>{isEn ? "Simulate text, formula and paragraph spacing according to the final import level" : "Mô phỏng khoảng cách văn bản, công thức và đoạn văn theo cấp độ nhập cuối cùng"}</small>
              </div>
            </div>
            {preview.parsed ? (
              <div className="doc-import-preview-counts" aria-label={isEn ? "Preview statistics" : "Thống kê xem trước"}>
                <span>{preview.parsed.blocks.length} {isEn ? "block" : ""}</span>
                <span>{preview.parsed.inlineFormulaCount} {isEn ? "inline" : "nội tuyến"}</span>
                <span>{preview.parsed.displayFormulaCount} {isEn ? "between lines" : "giữa các dòng"}</span>
              </div>
            ) : (
              <span className="doc-import-pane-stat">{isEn ? "Wait for valid content" : "Đợi nội dung hợp lệ"}</span>
            )}
          </div>
          <div className="doc-import-preview-scroll">
            {preview.parsed ? <PreviewPane parsed={preview.parsed} /> : (
              <div className="doc-import-preview-error">
                <TriangleAlert size={20} />
                <span>{preview.error}</span>
              </div>
            )}
          </div>
        </article>
      </section>

      <footer className="doc-import-footer">
        <div className="doc-import-messages">
          {loadError ? <span className="error"><TriangleAlert size={15} />{loadError}</span> : null}
          {preview.parsed?.warnings.map((warning, index) => (
            <span className="warning" key={index}><TriangleAlert size={15} />{warning}</span>
          ))}
          {!loadError && preview.parsed && preview.parsed.warnings.length === 0 ? (
            <span className="ok">
              <CheckCircle2 size={15} />
              {isEn ? "The preview parsing is normal; Word will insert according to the current structured preview results." : "Việc phân tích cú pháp xem trước là bình thường; Word sẽ chèn theo kết quả xem trước có cấu trúc hiện tại."}</span>
          ) : null}
        </div>
        <div className="doc-import-actions">
          <button className="doc-import-secondary" onClick={() => void cancel()} disabled={busy}>
            <X size={16} />{isEn ? "Cancel" : "Hủy bỏ"}</button>
          <button className="doc-import-primary" onClick={() => void commit()} disabled={busy || !preview.parsed || Boolean(preview.error)}>
            {busy ? <LoaderCircle size={16} className="spin" /> : <FileText size={16} />}
            {busy ? (isEn ? "Submitting…" : "Đang gửi…") : (isEn ? "Import into Word" : "Nhập vào Word")}
          </button>
        </div>
      </footer>
    </main>
  );
}
