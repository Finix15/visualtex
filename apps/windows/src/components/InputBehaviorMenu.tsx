import {
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  type CSSProperties,
} from "react";
import { createPortal } from "react-dom";
import { convertVisualTexLatexToMarkup } from "../editor/mathLiveIntegralCompatibility";
import {
  ArrowLeft,
  ArrowRight,
  ChevronDown,
  MousePointerClick,
  Search,
} from "lucide-react";
import type { InputBehaviorSettingKey } from "../types/formula";
import { useEditorStore } from "../stores/editorStore";
import {
  visualTexAutoEscapeInlineShortcuts,
  visualTexAutoEscapeShortcutGroups,
  type VisualTexInlineShortcutDefinition,
  type VisualTexInlineShortcutDefinitions,
} from "../editor/normalizeChineseLatex";

interface InputBehaviorOption {
  key: InputBehaviorSettingKey;
  titleVi: string;
  titleEn: string;
  descriptionVi: string;
  descriptionEn: string;
}

const AUTO_ESCAPE_OPTIONS: InputBehaviorOption[] = [
  {
    key: "autoEscapeShortcuts",
    titleVi: "Các phím tắt toán học thông dụng",
    titleEn: "Common math shortcuts",
    descriptionVi: "Điều khiển các phím tắt như alpha, >= và hat; phát hiện thẳng đứng cho sự khác biệt và tên hàm vẫn độc lập",
    descriptionEn: "Controls shortcuts such as alpha, >= and hat; upright detection for differentials and function names remains independent",
  },
];

const CARET_BEHAVIOR_OPTIONS: InputBehaviorOption[] = [
  {
    key: "autoExitSuperscript",
    titleVi: "Thoát chỉ số trên sau khi nhập",
    titleEn: "Exit superscript after input",
    descriptionVi: "Trở về công thức chính sau một ký tự hoặc ký hiệu trên thanh công cụ",
    descriptionEn: "Return to the main formula after one character or toolbar symbol",
  },
  {
    key: "autoExitSubscript",
    titleVi: "Thoát chỉ số sau khi nhập",
    titleEn: "Exit subscript after input",
    descriptionVi: "Trở về công thức chính sau một ký tự hoặc ký hiệu trên thanh công cụ",
    descriptionEn: "Return to the main formula after one character or toolbar symbol",
  },
  {
    key: "autoExitAccent",
    titleVi: "Thoát dấu sau khi nhập",
    titleEn: "Exit accent after input",
    descriptionVi: "Áp dụng cho các dấu hat, bar, vec, dấu ngã, dấu chấm và các dấu tương tự",
    descriptionEn: "Applies to hat, bar, vec, tilde, dot and similar accents",
  },
  {
    key: "autoExitWrapperCommand",
    titleVi: "Thoát lệnh phông chữ sau khi nhập",
    titleEn: "Exit font command after input",
    descriptionVi: "Khi được bật, thoát sau một ký tự; khi bị tắt, hãy tiếp tục gõ và nhấn Enter để xác nhận phạm vi phông chữ mathbb, mathbf, mathcal và các phông chữ tương tự",
    descriptionEn: "When enabled, exit after one character; when disabled, keep typing and press Enter to confirm mathbb, mathbf, mathcal and similar font scopes",
  },
];

const COMMAND_SUGGESTION_OPTIONS: InputBehaviorOption[] = [
  {
    key: "showStructuredCommandSuggestions",
    titleVi: "Gợi ý lệnh có cấu trúc",
    titleEn: "Structured command suggestions",
    descriptionVi: "Điều khiển bảng VisualTeX lớn để tính tổng, tích phân và các cấu trúc tương tự; không ảnh hưởng đến bảng lệnh gốc của MathLive",
    descriptionEn: "Controls the large VisualTeX panel for sums, integrals and similar structures; does not affect MathLive's native command panel",
  },
  {
    key: "showOtherCommandSuggestions",
    titleVi: "Gợi ý lệnh khác",
    titleEn: "Other command suggestions",
    descriptionVi: "Điều khiển bảng VisualTeX lớn cho các lệnh không phải là tổng, tích phân và các cấu trúc tương tự; tắt theo mặc định",
    descriptionEn: "Controls the large VisualTeX panel for commands other than sums, integrals and similar structures; off by default",
  },
];

function shortcutLatex(definition: VisualTexInlineShortcutDefinition) {
  return typeof definition === "string" ? definition : definition.value;
}

function shortcutAfter(definition: VisualTexInlineShortcutDefinition) {
  return typeof definition === "string" ? "" : definition.after ?? "";
}

function previewLatex(definition: VisualTexInlineShortcutDefinition) {
  return shortcutLatex(definition).replaceAll("#?", "\\square");
}

function readActiveInlineShortcuts(): VisualTexInlineShortcutDefinitions {
  try {
    const field = document.querySelector("math-field") as
      | (HTMLElement & {
          inlineShortcuts?: Readonly<VisualTexInlineShortcutDefinitions>;
        })
      | null;
    const active = field?.inlineShortcuts;
    if (active && Object.keys(active).length > 0) return { ...active };
  } catch {
    // Fall back to the explicit VisualTeX table while the mathfield mounts.
  }
  return { ...visualTexAutoEscapeInlineShortcuts };
}

interface InputBehaviorPopoverLayout {
  left: number;
  top: number;
  width: number;
  maxHeight: number;
  compact: boolean;
}

function ShortcutOutput({ definition }: { definition: VisualTexInlineShortcutDefinition }) {
  const latex = previewLatex(definition);
  const markup = useMemo(() => {
    try {
      return convertVisualTexLatexToMarkup(latex, { defaultMode: "math" });
    } catch {
      return "";
    }
  }, [latex]);

  if (!markup) return <code>{shortcutLatex(definition)}</code>;
  return (
    <span
      className="auto-escape-map-output-formula"
      dangerouslySetInnerHTML={{ __html: markup }}
    />
  );
}

export function InputBehaviorMenu() {
  const [open, setOpen] = useState(false);
  const [page, setPage] = useState<"settings" | "mappings">("settings");
  const [mappingQuery, setMappingQuery] = useState("");
  const [activeShortcutDefinitions, setActiveShortcutDefinitions] =
    useState<VisualTexInlineShortcutDefinitions>(() => ({
      ...visualTexAutoEscapeInlineShortcuts,
    }));
  const rootRef = useRef<HTMLDivElement>(null);
  const triggerRef = useRef<HTMLButtonElement>(null);
  const popoverRef = useRef<HTMLDivElement>(null);
  const [popoverLayout, setPopoverLayout] =
    useState<InputBehaviorPopoverLayout | null>(null);
  const language = useEditorStore((state) => state.language);
  const inputBehavior = useEditorStore((state) => state.inputBehavior);
  const setInputBehavior = useEditorStore((state) => state.setInputBehavior);
  const isEn = language === "en";

  const filteredShortcutGroups = useMemo(() => {
    const query = mappingQuery.trim().toLocaleLowerCase();
    const seen = new Set<string>();
    const groups = visualTexAutoEscapeShortcutGroups.map((group) => {
      const entries = Object.keys(group.shortcuts)
        .filter((shortcut) => shortcut in activeShortcutDefinitions)
        .map(
          (shortcut) =>
            [shortcut, activeShortcutDefinitions[shortcut]] as const,
        );
      entries.forEach(([shortcut]) => seen.add(shortcut));
      return { ...group, entries };
    });
    const mathLiveEntries = Object.entries(activeShortcutDefinitions).filter(
      ([shortcut]) => !seen.has(shortcut),
    );
    if (mathLiveEntries.length > 0) {
      groups.push({
        id: "mathlive",
        titleVi: "Tích hợp sẵn MathLive",
        titleEn: "MathLive built-ins",
        shortcuts: {},
        entries: mathLiveEntries,
      });
    }
    return groups
      .map((group) => ({
        ...group,
        entries: group.entries.filter(([shortcut, definition]) => {
          if (!query) return true;
          return (
            shortcut.toLocaleLowerCase().includes(query) ||
            shortcutLatex(definition).toLocaleLowerCase().includes(query)
          );
        }),
      }))
      .filter((group) => group.entries.length > 0);
  }, [activeShortcutDefinitions, mappingQuery]);

  const mappingCount = useMemo(
    () => filteredShortcutGroups.reduce((sum, group) => sum + group.entries.length, 0),
    [filteredShortcutGroups],
  );

  useEffect(() => {
    if (!open) return;
    const handlePointerDown = (event: PointerEvent) => {
      const target = event.target as Node;
      if (
        !rootRef.current?.contains(target) &&
        !popoverRef.current?.contains(target)
      ) {
        setOpen(false);
      }
    };
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key !== "Escape") return;
      if (page === "mappings") {
        setPage("settings");
        setMappingQuery("");
      } else {
        setOpen(false);
      }
    };
    document.addEventListener("pointerdown", handlePointerDown, true);
    document.addEventListener("keydown", handleKeyDown);
    return () => {
      document.removeEventListener("pointerdown", handlePointerDown, true);
      document.removeEventListener("keydown", handleKeyDown);
    };
  }, [open, page]);

  useEffect(() => {
    if (open) return;
    setPage("settings");
    setMappingQuery("");
    setPopoverLayout(null);
  }, [open]);

  useLayoutEffect(() => {
    if (!open) return;
    const trigger = triggerRef.current;
    if (!trigger) return;

    let frame = 0;
    const updateLayout = () => {
      window.cancelAnimationFrame(frame);
      frame = window.requestAnimationFrame(() => {
        const triggerRect = trigger.getBoundingClientRect();
        const workspace = trigger.closest<HTMLElement>(".workspace");
        const visibleEditor = workspace?.classList.contains("is-classic-layout")
          ? workspace.querySelector<HTMLElement>(".classic-editor-pane-body")
          : workspace?.querySelector<HTMLElement>(".formula-workspace.editor-pane");
        const editorRect = visibleEditor?.getBoundingClientRect();
        const viewportMargin = 8;
        const popoverGap = 6;
        const viewportRight = Math.max(viewportMargin, window.innerWidth - viewportMargin);
        const viewportBottom = Math.max(viewportMargin, window.innerHeight - viewportMargin);
        const editorIsUsable = Boolean(
          editorRect && editorRect.width >= 140 && editorRect.height >= 100,
        );
        const leftBound = editorIsUsable
          ? Math.max(viewportMargin, editorRect!.left + viewportMargin)
          : viewportMargin;
        const rightBound = editorIsUsable
          ? Math.min(viewportRight, editorRect!.right - viewportMargin)
          : viewportRight;
        const availableWidth = Math.max(120, rightBound - leftBound);
        const preferredWidth = page === "mappings" ? 660 : 420;
        const width = Math.min(preferredWidth, availableWidth);
        const left = Math.min(
          Math.max(triggerRect.right - width, leftBound),
          Math.max(leftBound, rightBound - width),
        );
        const minimumTop = editorIsUsable
          ? Math.max(viewportMargin, editorRect!.top + viewportMargin)
          : viewportMargin;
        const top = Math.min(
          Math.max(triggerRect.bottom + popoverGap, minimumTop),
          Math.max(viewportMargin, viewportBottom - 96),
        );
        const preferredMaxHeight = page === "mappings" ? 620 : 560;
        const maxHeight = Math.max(
          96,
          Math.min(preferredMaxHeight, viewportBottom - top),
        );
        const next = {
          left,
          top,
          width,
          maxHeight,
          compact: width < 360 || maxHeight < 420,
        };
        setPopoverLayout((current) =>
          current &&
          Math.abs(current.left - next.left) < 0.5 &&
          Math.abs(current.top - next.top) < 0.5 &&
          Math.abs(current.width - next.width) < 0.5 &&
          Math.abs(current.maxHeight - next.maxHeight) < 0.5 &&
          current.compact === next.compact
            ? current
            : next,
        );
      });
    };

    const resizeObserver = new ResizeObserver(updateLayout);
    resizeObserver.observe(trigger);
    const workspace = trigger.closest<HTMLElement>(".workspace");
    if (workspace) resizeObserver.observe(workspace);
    window.addEventListener("resize", updateLayout);
    window.addEventListener("scroll", updateLayout, true);
    updateLayout();
    return () => {
      window.cancelAnimationFrame(frame);
      resizeObserver.disconnect();
      window.removeEventListener("resize", updateLayout);
      window.removeEventListener("scroll", updateLayout, true);
    };
  }, [open, page]);

  const popoverStyle = popoverLayout
    ? ({
        left: `${popoverLayout.left}px`,
        top: `${popoverLayout.top}px`,
        width: `${popoverLayout.width}px`,
        maxWidth: `${popoverLayout.width}px`,
        maxHeight: `${popoverLayout.maxHeight}px`,
      } as CSSProperties)
    : undefined;

  return (
    <div ref={rootRef} className="input-behavior-menu">
      <button
        ref={triggerRef}
        type="button"
        className={`canvas-input-behavior-trigger${open ? " is-active" : ""}`}
        aria-haspopup="dialog"
        aria-expanded={open}
        onClick={() => setOpen((value) => !value)}
        title={isEn ? "Input behavior" : "Hành vi nhập liệu"}
      >
        <MousePointerClick size={14} />
        <span>{isEn ? "Input behavior" : "Hành vi nhập liệu"}</span>
        <ChevronDown size={13} aria-hidden="true" />
      </button>

      {open &&
        popoverLayout &&
        createPortal(
        <div
          ref={popoverRef}
          className={
            `input-behavior-popover${page === "mappings" ? " is-mapping-view" : ""}` +
            (popoverLayout.compact ? " is-compact" : "")
          }
          style={popoverStyle}
          role="dialog"
          aria-label={
            page === "mappings"
              ? isEn
                ? "Automatic conversion mappings"
                : "Ánh xạ chuyển đổi tự động"
              : isEn
                ? "Input behavior settings"
                : "Cài đặt hành vi nhập liệu"
          }
        >
          {page === "mappings" ? (
            <>
              <div className="auto-escape-map-toolbar">
                <button
                  type="button"
                  className="auto-escape-map-back"
                  aria-label={isEn ? "Back to input behavior" : "Quay lại thao tác nhập liệu"}
                  onClick={() => {
                    setPage("settings");
                    setMappingQuery("");
                  }}
                >
                  <ArrowLeft size={15} />
                </button>
                <div>
                  <strong>{isEn ? "Conversion mappings" : "Ánh xạ chuyển đổi"}</strong>
                  <span>{mappingCount}</span>
                </div>
              </div>

              <label className="auto-escape-map-search">
                <Search size={14} aria-hidden="true" />
                <input
                  value={mappingQuery}
                  onChange={(event) => setMappingQuery(event.target.value)}
                  placeholder={isEn ? "Search input or LaTeX" : "Đầu vào tìm kiếm hoặc LaTeX"}
                  aria-label={isEn ? "Search conversion mappings" : "Ánh xạ chuyển đổi tìm kiếm"}
                  autoFocus
                />
              </label>

              <div className="auto-escape-map-groups">
                {filteredShortcutGroups.map((group) => (
                  <section className="auto-escape-map-group" key={group.id}>
                    <div className="auto-escape-map-group-heading">
                      <strong>{isEn ? group.titleEn : group.titleVi}</strong>
                      {group.entries.some(([, definition]) => shortcutAfter(definition)) ? (
                        <span
                          title={
                            isEn
                              ? "These entries apply only after the allowed preceding structures defined in code"
                              : "Các mục này chỉ áp dụng sau các cấu trúc được phép trước đó được xác định trong mã"
                          }
                        >
                          {isEn ? "context" : "bối cảnh"}
                        </span>
                      ) : null}
                    </div>
                    <div className="auto-escape-map-grid">
                      {group.entries.map(([shortcut, definition]) => (
                        <div
                          className="auto-escape-map-row"
                          key={shortcut}
                          data-auto-escape-shortcut={shortcut}
                          data-auto-escape-output={shortcutLatex(definition)}
                          data-auto-escape-after={shortcutAfter(definition)}
                          title={`${shortcut} → ${shortcutLatex(definition)}${
                            shortcutAfter(definition)
                              ? ` · after: ${shortcutAfter(definition)}`
                              : ""
                          }`}
                        >
                          <span className="auto-escape-map-input-wrap">
                            <code className="auto-escape-map-input">{shortcut}</code>
                            {shortcutAfter(definition) ? (
                              <span
                                className="auto-escape-condition-mark"
                                aria-label={isEn ? "Has context condition" : "Có điều kiện ngữ cảnh"}
                              />
                            ) : null}
                          </span>
                          <ArrowRight size={13} aria-hidden="true" />
                          <span className="auto-escape-map-output">
                            <ShortcutOutput definition={definition} />
                          </span>
                        </div>
                      ))}
                    </div>
                  </section>
                ))}
                {filteredShortcutGroups.length === 0 ? (
                  <div className="auto-escape-map-empty">
                    {isEn ? "No matching mapping" : "Không có ánh xạ phù hợp"}
                  </div>
                ) : null}
              </div>
            </>
          ) : (
            <>
              <div className="input-behavior-heading">
                <strong>{isEn ? "Automatic conversion" : "Chuyển đổi tự động"}</strong>
              </div>

              <div className="input-behavior-options">
                {AUTO_ESCAPE_OPTIONS.map((option) => (
                  <div className="input-behavior-option has-secondary-action" key={option.key}>
                    <span>
                      <strong>{isEn ? option.titleEn : option.titleVi}</strong>
                      <button
                        type="button"
                        className="input-behavior-map-button"
                        data-open-auto-escape-map
                        onClick={() => {
                          setActiveShortcutDefinitions(readActiveInlineShortcuts());
                          setPage("mappings");
                        }}
                      >
                        {isEn ? "View mappings" : "Xem bản đồ"}
                        <ArrowRight size={13} aria-hidden="true" />
                      </button>
                    </span>
                    <label
                      className="input-behavior-toggle"
                      aria-label={isEn ? option.titleEn : option.titleVi}
                    >
                      <input
                        type="checkbox"
                        checked={inputBehavior[option.key]}
                        onChange={(event) =>
                          setInputBehavior(option.key, event.target.checked)
                        }
                      />
                      <span className="input-behavior-switch" aria-hidden="true" />
                    </label>
                  </div>
                ))}
              </div>

              <div className="input-behavior-heading input-behavior-section-heading">
                <strong>{isEn ? "Caret auto-exit" : "Tự động thoát dấu mũ"}</strong>
              </div>

              <div className="input-behavior-options">
                {CARET_BEHAVIOR_OPTIONS.map((option) => (
                  <label className="input-behavior-option" key={option.key}>
                    <span>
                      <strong>{isEn ? option.titleEn : option.titleVi}</strong>
                    </span>
                    <input
                      type="checkbox"
                      checked={inputBehavior[option.key]}
                      onChange={(event) => setInputBehavior(option.key, event.target.checked)}
                    />
                    <span className="input-behavior-switch" aria-hidden="true" />
                  </label>
                ))}
              </div>

              <div className="input-behavior-heading input-behavior-section-heading">
                <strong>{isEn ? "Command suggestion panels" : "Bảng gợi ý lệnh"}</strong>
              </div>

              <div className="input-behavior-options">
                {COMMAND_SUGGESTION_OPTIONS.map((option) => (
                  <label className="input-behavior-option" key={option.key}>
                    <span>
                      <strong>{isEn ? option.titleEn : option.titleVi}</strong>
                    </span>
                    <input
                      type="checkbox"
                      checked={inputBehavior[option.key]}
                      onChange={(event) => setInputBehavior(option.key, event.target.checked)}
                    />
                    <span className="input-behavior-switch" aria-hidden="true" />
                  </label>
                ))}
              </div>
            </>
          )}
        </div>,
        document.body,
      )}
    </div>
  );
}
