import { useEffect, useRef } from "react";
import { Clock3, Trash2, X } from "lucide-react";
import { MathPreview } from "./MathPreview";
import { useEditorStore } from "../stores/editorStore";

interface Props {
  open: boolean;
  onClose: () => void;
  onRestore: (latex: string) => void;
}

export function HistoryPanel({ open, onClose, onRestore }: Props) {
  const panelRef = useRef<HTMLElement>(null);
  const previousFocusRef = useRef<HTMLElement | null>(null);
  const history = useEditorStore((state) => state.history);
  const clearHistory = useEditorStore((state) => state.clearHistory);
  const language = useEditorStore((state) => state.language);
  const isEn = language === "en";

  useEffect(() => {
    if (!open) return;
    previousFocusRef.current = document.activeElement as HTMLElement | null;
    const frame = window.requestAnimationFrame(() => {
      panelRef.current?.querySelector<HTMLElement>("button")?.focus();
    });
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.preventDefault();
        onClose();
      }
    };
    document.addEventListener("keydown", handleKeyDown);
    return () => {
      window.cancelAnimationFrame(frame);
      document.removeEventListener("keydown", handleKeyDown);
      previousFocusRef.current?.focus({ preventScroll: true });
    };
  }, [open]);

  const timeLabel = (time: number) =>
    new Intl.DateTimeFormat(isEn ? "en-US" : "vi-VN", {
      month: "short",
      day: "numeric",
      hour: "2-digit",
      minute: "2-digit",
    }).format(time);

  if (!open) return null;

  return (
    <aside
      ref={panelRef}
      className="history-panel is-open"
      role="dialog"
      aria-modal="true"
      aria-labelledby="history-panel-title"
    >
      <header className="history-header">
        <div>
          <span className="eyebrow">RECENT</span>
          <h2 id="history-panel-title">{isEn ? "Formula history" : "Lịch sử công thức"}</h2>
        </div>
        <button
          type="button"
          className="icon-button"
          onClick={onClose}
          aria-label={isEn ? "Close history" : "Đóng lịch sử"}
        >
          <X size={18} />
        </button>
      </header>

      <div className="history-list">
        {history.length === 0 ? (
          <div className="empty-state">
            <span className="empty-icon"><Clock3 size={24} /></span>
            <h3>{isEn ? "No formula history yet" : "Chưa có lịch sử công thức"}</h3>
            <p>
              {isEn
                ? "The current formula will be saved here after you pause editing."
                : "Công thức hiện tại sẽ được lưu tại đây sau khi bạn tạm dừng chỉnh sửa."}
            </p>
          </div>
        ) : (
          history.map((item) => {
            const itemLines = item.latex.split("\n").filter((line) => line.trim());
            return (
              <button
                type="button"
                className="history-item"
                key={item.id}
                onClick={() => onRestore(item.latex)}
              >
                <span className="history-formula-stack">
                  {itemLines.slice(0, 3).map((line, index) => (
                    <MathPreview latex={line} key={index} />
                  ))}
                  {itemLines.length > 3 && (
                    <small>+{itemLines.length - 3} {isEn ? "lines" : "dòng"}</small>
                  )}
                </span>
                <span>{timeLabel(item.createdAt)}</span>
              </button>
            );
          })
        )}
      </div>

      {history.length > 0 && (
        <footer className="history-footer">
          <button type="button" className="text-button danger-text" onClick={clearHistory}>
            <Trash2 size={14} /> {isEn ? "Clear history" : "Xóa lịch sử"}
          </button>
        </footer>
      )}
    </aside>
  );
}
