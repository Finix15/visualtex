import {
  useEffect,
  useRef,
  useState,
  type ChangeEvent,
  type CSSProperties,
} from "react";
import { createPortal } from "react-dom";
import { invoke, isTauri } from "@tauri-apps/api/core";
import { save } from "@tauri-apps/plugin-dialog";
import {
  BrainCircuit,
  ChevronRight,
  Download,
  Eye,
  Keyboard,
  Languages,
  Presentation,
  RefreshCw,
  RotateCcw,
  SlidersHorizontal,
  Upload,
  X,
} from "lucide-react";
import {
  DEFAULT_FORMULA_INSET,
  DEFAULT_FORMULA_ROW_VERTICAL_INSET,
  DEFAULT_FORMULA_TOOL_BUTTON_PADDING,
  DEFAULT_FORMULA_TOOL_BUTTON_SIZE,
  EDITOR_ZOOM_STEP,
  MAX_FORMULA_INSET,
  MAX_FORMULA_ROW_VERTICAL_INSET,
  MAX_FORMULA_TOOL_BUTTON_PADDING,
  MAX_FORMULA_TOOL_BUTTON_SIZE,
  MIN_FORMULA_INSET,
  MIN_FORMULA_ROW_VERTICAL_INSET,
  MIN_FORMULA_TOOL_BUTTON_PADDING,
  MIN_FORMULA_TOOL_BUTTON_SIZE,
  useEditorStore,
} from "../stores/editorStore";
import { MathPreview } from "./MathPreview";
import { OfficeIntegrationSettings } from "./OfficeIntegrationSettings";
import { pngExportBackgroundPickerValue } from "../export/pngBackground";
import {
  DEFAULT_FORMULA_CHINESE_FONT,
  DEFAULT_FORMULA_LETTER_FONT,
  FORMULA_CHINESE_FONT_OPTIONS,
  FORMULA_LETTER_FONT_OPTIONS,
  formulaChineseFontFamily,
  formulaLetterFontFamilies,
  type FormulaChineseFont,
  type FormulaLetterFont,
} from "../editor/formulaFontPreferences";
import {
  applyVisualTexConfiguration,
  buildVisualTexConfiguration,
  parseVisualTexConfiguration,
  VISUALTEX_CONFIGURATION_EXTENSION,
} from "../runtime/applicationConfiguration";
import {
  createDefaultCustomTheme,
  publishCustomTheme,
  readCustomTheme,
  THEME_DEFINITIONS,
  type CustomThemeState,
  type ThemePaletteColors,
} from "../themeCustomization";
import { publishSynchronizedTheme } from "../themeSync";

const THEME_CORE_COLOR_FIELDS: readonly [keyof ThemePaletteColors, string, string][] = [
  ["accent", "Accent", "Giọng"],
  ["accentHover", "Accent hover", "Di chuột có dấu"],
  ["accentSoft", "Accent soft", "Giọng nhẹ nhàng"],
  ["background", "App background", "Nền ứng dụng"],
  ["elevated", "Raised background", "Nền nâng"],
  ["surface", "Panel surface", "Bề mặt tấm"],
  ["sunken", "Sunken surface", "Bề mặt trũng"],
  ["hover", "Hover surface", "Bề mặt di chuột"],
  ["active", "Selected surface", "Bề mặt được chọn"],
  ["foreground", "Primary text", "Văn bản chính"],
  ["textMuted", "Secondary text", "Văn bản phụ"],
  ["textFaint", "Muted text", "Văn bản bị tắt tiếng"],
  ["border", "Border", "Viền"],
  ["borderStrong", "Strong border", "Viền chắc chắn"],
  ["focusRing", "Focus ring", "Vòng lấy nét"],
  ["info", "Info", "Thông tin"],
  ["success", "Success", "Thành công"],
  ["warning", "Warning", "Cảnh báo"],
  ["danger", "Danger", "Nguy hiểm"],
];

const THEME_FORMULA_COLOR_FIELDS: readonly [keyof ThemePaletteColors, string, string][] = [
  ["formulaSurface", "Formula canvas", "Canvas công thức"],
  ["formulaPlaceholder", "Placeholder", "Phần giữ chỗ"],
  ["formulaPlaceholderSelected", "Selected placeholder", "Phần giữ chỗ đã chọn"],
  ["formulaCaret", "Formula caret", "Dấu mũ công thức"],
];

const THEME_SYNTAX_COLOR_FIELDS: readonly [keyof ThemePaletteColors, string, string][] = [
  ["syntaxCommand", "Command", "Lệnh"],
  ["syntaxKeyword", "Keyword", "Từ khóa"],
  ["syntaxOperator", "Operator", "Người vận hành"],
  ["syntaxNumber", "Number", "Số"],
  ["syntaxBracket", "Bracket", "Giá đỡ"],
  ["syntaxString", "String", "Chuỗi"],
  ["syntaxComment", "Comment", "Bình luận"],
  ["syntaxVariable", "Variable", "Biến"],
  ["syntaxFunction", "Function", "Chức năng"],
  ["syntaxError", "Error", "Lỗi"],
];

const THEME_TOOLBAR_COLOR_FIELDS: readonly [keyof ThemePaletteColors, string, string][] = [
  ["toolbarCommon", "Common", "Chung"],
  ["toolbarStructure", "Structure", "Cấu trúc"],
  ["toolbarCalculus", "Calculus", "Giải tích"],
  ["toolbarMatrix", "Matrix", "Ma trận"],
  ["toolbarRelation", "Relation", "Quan hệ"],
  ["toolbarGreek", "Greek", "Tiếng Hy Lạp"],
  ["toolbarArrow", "Arrow", "Mũi tên"],
  ["toolbarPhysics", "Physics", "Vật Lý"],
  ["toolbarSet", "Set", "Bộ"],
];

interface Props {
  open: boolean;
  onClose: () => void;
  onCheckForUpdates: () => void;
  onOpenFormulaHotkeys: () => void;
}

function encodeUtf8Base64(value: string) {
  const bytes = new TextEncoder().encode(value);
  let binary = "";
  const chunkSize = 0x8000;
  for (let offset = 0; offset < bytes.length; offset += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + chunkSize));
  }
  return btoa(binary);
}

function configurationFilename() {
  return `VisualTeX-Configuration-${new Date().toISOString().slice(0, 10)}.${VISUALTEX_CONFIGURATION_EXTENSION}`;
}

function downloadConfigurationInBrowser(contents: string, filename: string) {
  const blob = new Blob([contents], { type: "application/json;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 0);
}

export function SettingsDialog({
  open,
  onClose,
  onCheckForUpdates,
  onOpenFormulaHotkeys,
}: Props) {
  const dialogRef = useRef<HTMLElement>(null);
  const interfaceCustomizationDialogRef = useRef<HTMLElement>(null);
  const interfaceCustomizationOpenRef = useRef(false);
  const previousFocusRef = useRef<HTMLElement | null>(null);
  const configurationInputRef = useRef<HTMLInputElement>(null);
  const [interfaceCustomizationOpen, setInterfaceCustomizationOpen] =
    useState(false);
  const [configurationBusy, setConfigurationBusy] = useState(false);
  const [configurationStatus, setConfigurationStatus] = useState("");
  const [customTheme, setCustomTheme] =
    useState<CustomThemeState>(() => readCustomTheme());
  const theme = useEditorStore((state) => state.theme);
  const setTheme = useEditorStore((state) => state.setTheme);
  const language = useEditorStore((state) => state.language);
  const setLanguage = useEditorStore((state) => state.setLanguage);
  const zoom = useEditorStore((state) => state.zoom);
  const setZoom = useEditorStore((state) => state.setZoom);
  const editorLayout = useEditorStore((state) => state.editorLayout);
  const setEditorLayout = useEditorStore((state) => state.setEditorLayout);
  const autoPairDelimiters = useEditorStore(
    (state) => state.autoPairDelimiters,
  );
  const showLineNumbers = useEditorStore((state) => state.showLineNumbers);
  const highlightActiveLine = useEditorStore(
    (state) => state.highlightActiveLine,
  );
  const keypadMinimizeOnCopy = useEditorStore(
    (state) => state.keypadMinimizeOnCopy,
  );
  const formulaInsetLeft = useEditorStore((state) => state.formulaInsetLeft);
  const formulaInsetRight = useEditorStore((state) => state.formulaInsetRight);
  const formulaToolButtonSize = useEditorStore(
    (state) => state.formulaToolButtonSize,
  );
  const formulaToolButtonPadding = useEditorStore(
    (state) => state.formulaToolButtonPadding,
  );
  const formulaRowVerticalInset = useEditorStore(
    (state) => state.formulaRowVerticalInset,
  );
  const pngExportBackground = useEditorStore(
    (state) => state.pngExportBackground,
  );
  const formulaLetterFont = useEditorStore((state) => state.formulaLetterFont);
  const formulaChineseFont = useEditorStore((state) => state.formulaChineseFont);
  const setAutoPairDelimiters = useEditorStore(
    (state) => state.setAutoPairDelimiters,
  );
  const setShowLineNumbers = useEditorStore(
    (state) => state.setShowLineNumbers,
  );
  const setHighlightActiveLine = useEditorStore(
    (state) => state.setHighlightActiveLine,
  );
  const setKeypadMinimizeOnCopy = useEditorStore(
    (state) => state.setKeypadMinimizeOnCopy,
  );
  const setFormulaInsetLeft = useEditorStore(
    (state) => state.setFormulaInsetLeft,
  );
  const setFormulaInsetRight = useEditorStore(
    (state) => state.setFormulaInsetRight,
  );
  const setFormulaToolButtonSize = useEditorStore(
    (state) => state.setFormulaToolButtonSize,
  );
  const setFormulaToolButtonPadding = useEditorStore(
    (state) => state.setFormulaToolButtonPadding,
  );
  const setFormulaRowVerticalInset = useEditorStore(
    (state) => state.setFormulaRowVerticalInset,
  );
  const setPngExportBackground = useEditorStore(
    (state) => state.setPngExportBackground,
  );
  const setFormulaLetterFont = useEditorStore(
    (state) => state.setFormulaLetterFont,
  );
  const setFormulaChineseFont = useEditorStore(
    (state) => state.setFormulaChineseFont,
  );
  const personalize = useEditorStore((state) => state.personalize);
  const setPersonalize = useEditorStore((state) => state.setPersonalize);
  const suggestionCount = useEditorStore((state) => state.suggestionCount);
  const setSuggestionCount = useEditorStore((state) => state.setSuggestionCount);
  const resetUsage = useEditorStore((state) => state.resetUsage);
  const powerPointDefaultFontSizePt = useEditorStore(
    (state) => state.powerPointDefaultFontSizePt,
  );
  const setPowerPointDefaultFontSizePt = useEditorStore(
    (state) => state.setPowerPointDefaultFontSizePt,
  );
  const checkUpdatesOnStartup = useEditorStore(
    (state) => state.checkUpdatesOnStartup,
  );
  const setCheckUpdatesOnStartup = useEditorStore(
    (state) => state.setCheckUpdatesOnStartup,
  );
  const isEn = language === "en";
  const selectTheme = (nextTheme: Parameters<typeof setTheme>[0]) => {
    setTheme(nextTheme);
    publishSynchronizedTheme(nextTheme);
  };
  const formulaLetterFamilies = formulaLetterFontFamilies(formulaLetterFont);
  const formulaChineseFamily = formulaChineseFontFamily(formulaChineseFont);

  useEffect(() => {
    if (!open) return;
    previousFocusRef.current = document.activeElement as HTMLElement | null;
    const frame = window.requestAnimationFrame(() => {
      dialogRef.current?.querySelector<HTMLElement>("button, input, select")?.focus();
    });

    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.preventDefault();
        if (interfaceCustomizationOpenRef.current) {
          interfaceCustomizationOpenRef.current = false;
          setInterfaceCustomizationOpen(false);
        } else {
          onClose();
        }
        return;
      }
      const activeDialog = interfaceCustomizationOpenRef.current
        ? interfaceCustomizationDialogRef.current
        : dialogRef.current;
      if (event.key !== "Tab" || !activeDialog) return;

      const focusable = Array.from(
        activeDialog.querySelectorAll<HTMLElement>(
          'button:not(:disabled), input:not(:disabled), select:not(:disabled), [tabindex]:not([tabindex="-1"])',
        ),
      );
      const first = focusable[0];
      const last = focusable[focusable.length - 1];
      if (!first || !last) return;

      if (event.shiftKey && document.activeElement === first) {
        event.preventDefault();
        last.focus();
      } else if (!event.shiftKey && document.activeElement === last) {
        event.preventDefault();
        first.focus();
      }
    };

    document.addEventListener("keydown", handleKeyDown);
    return () => {
      window.cancelAnimationFrame(frame);
      document.removeEventListener("keydown", handleKeyDown);
      previousFocusRef.current?.focus({ preventScroll: true });
    };
  }, [open]);

  useEffect(() => {
    interfaceCustomizationOpenRef.current = interfaceCustomizationOpen;
    if (!interfaceCustomizationOpen) return;
    const frame = window.requestAnimationFrame(() => {
      interfaceCustomizationDialogRef.current
        ?.querySelector<HTMLElement>("button, input")
        ?.focus();
    });
    return () => window.cancelAnimationFrame(frame);
  }, [interfaceCustomizationOpen]);

  useEffect(() => {
    if (open) return;
    interfaceCustomizationOpenRef.current = false;
    setInterfaceCustomizationOpen(false);
  }, [open]);

  const openInterfaceCustomization = () => {
    interfaceCustomizationOpenRef.current = true;
    setInterfaceCustomizationOpen(true);
  };
  const closeInterfaceCustomization = () => {
    interfaceCustomizationOpenRef.current = false;
    setInterfaceCustomizationOpen(false);
  };

  const updateCustomTheme = (
    update: (current: CustomThemeState) => CustomThemeState,
  ) => {
    setCustomTheme((current) => {
      const next = update(current);
      publishCustomTheme(next);
      selectTheme("custom");
      return next;
    });
  };

  const updateThemeColor = (key: keyof ThemePaletteColors, color: string) => {
    if (!/^#[0-9a-f]{6}$/i.test(color)) return;
    updateCustomTheme((current) => ({
      ...current,
      colors: {
        ...current.colors,
        [key]: color.toUpperCase(),
      },
    }));
  };

  const resetThemePalette = () => {
    const next = createDefaultCustomTheme();
    setCustomTheme(next);
    publishCustomTheme(next);
    selectTheme("custom");
  };

  const renderThemeColorFields = (
    fields: readonly [keyof ThemePaletteColors, string, string][],
  ) => (
    <div className="theme-color-grid">
      {fields.map(([key, labelEn, labelVi]) => {
        const color = customTheme.colors[key];
        return (
          <label className="theme-color-row" key={key}>
            <strong>{isEn ? labelEn : labelVi}</strong>
            <span className="theme-color-control">
              <input
                type="color"
                value={color}
                aria-label={isEn ? `${labelEn} color` : `${labelEn} màu`}
                data-theme-color-setting={key}
                onChange={(event) =>
                  updateThemeColor(key, event.currentTarget.value)
                }
              />
              <input
                key={`${key}-${color}`}
                type="text"
                defaultValue={color}
                spellCheck={false}
                aria-label={isEn ? `${labelEn} hex value` : `${labelEn} giá trị thập lục phân`}
                onBlur={(event) => {
                  const value = event.currentTarget.value.trim();
                  if (/^#[0-9a-f]{6}$/i.test(value)) {
                    updateThemeColor(key, value);
                  } else {
                    event.currentTarget.value = color;
                  }
                }}
                onKeyDown={(event) => {
                  if (event.key === "Enter") event.currentTarget.blur();
                }}
              />
            </span>
          </label>
        );
      })}
    </div>
  );

  const saveConfiguration = async () => {
    if (configurationBusy) return;
    setConfigurationBusy(true);
    setConfigurationStatus("");
    try {
      const configuration = await buildVisualTexConfiguration();
      const contents = JSON.stringify(configuration, null, 2);
      const filename = configurationFilename();
      if (isTauri()) {
        const selectedPath = await save({
          defaultPath: filename,
          filters: [
            {
              name: "VisualTeX Configuration",
              extensions: [VISUALTEX_CONFIGURATION_EXTENSION],
            },
          ],
        });
        if (!selectedPath) return;
        const path = selectedPath.toLowerCase().endsWith(`.${VISUALTEX_CONFIGURATION_EXTENSION}`)
          ? selectedPath
          : `${selectedPath}.${VISUALTEX_CONFIGURATION_EXTENSION}`;
        await invoke("write_export_file", {
          request: {
            path,
            base64: encodeUtf8Base64(contents),
          },
        });
      } else {
        downloadConfigurationInBrowser(contents, filename);
      }
      setConfigurationStatus(
        isEn ? "Configuration saved." : "Đã lưu cấu hình.",
      );
    } catch (reason) {
      const message =
        reason instanceof Error
          ? reason.message
          : typeof reason === "string"
            ? reason
            : "";
      setConfigurationStatus(
        message ||
          (isEn
            ? "Unable to save the configuration."
            : "Không thể lưu cấu hình."),
      );
    } finally {
      setConfigurationBusy(false);
    }
  };

  const importConfigurationFile = async (
    event: ChangeEvent<HTMLInputElement>,
  ) => {
    const file = event.currentTarget.files?.[0];
    event.currentTarget.value = "";
    if (!file || configurationBusy) return;
    setConfigurationBusy(true);
    setConfigurationStatus("");
    try {
      const configuration = parseVisualTexConfiguration(await file.text());
      await applyVisualTexConfiguration(configuration);
      setConfigurationStatus(
        isEn
          ? "Configuration imported. Reloading VisualTeX…"
          : "Đã nhập cấu hình. Đang tải lại VisualTeX…",
      );
      window.setTimeout(() => window.location.reload(), 450);
    } catch (reason) {
      setConfigurationStatus(
        reason instanceof Error
          ? reason.message
          : isEn
            ? "Unable to import the configuration."
            : "Không thể nhập cấu hình.",
      );
      setConfigurationBusy(false);
    }
  };

  if (!open) return null;

  return (
    <div className="modal-backdrop" role="presentation" onMouseDown={onClose}>
      <section
        ref={dialogRef}
        className="settings-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="settings-title"
        onMouseDown={(event) => event.stopPropagation()}
      >
        <header className="dialog-header">
          <div>
            <span className="eyebrow">PREFERENCES</span>
            <h2 id="settings-title">{isEn ? "Settings" : "Cài đặt"}</h2>
          </div>
          <button
            type="button"
            className="icon-button"
            onClick={onClose}
            aria-label={isEn ? "Close settings" : "Đóng cài đặt"}
          >
            <X size={18} />
          </button>
        </header>

        <div className="settings-content">
          <div className="settings-section">
            <div className="settings-section-title">
              <BrainCircuit size={18} />
              <div>
                <h3>{isEn ? "Personalized commands" : "Lệnh được cá nhân hóa"}</h3>
              </div>
            </div>
            <label className="switch-row">
              <span>
                <strong>{isEn ? "Enable personalized ranking" : "Kích hoạt tính năng xếp hạng được cá nhân hóa"}</strong>
              </span>
              <input
                type="checkbox"
                checked={personalize}
                onChange={(event) => setPersonalize(event.target.checked)}
              />
              <span className="switch-control" />
            </label>
            <label className="range-setting">
              <span>
                <strong>{isEn ? "Suggestion count" : "Số lượng đề xuất"}</strong>
                <small>
                  {suggestionCount} {isEn ? "items" : "mục"}
                </small>
              </span>
              <input
                type="range"
                min="3"
                max="10"
                value={suggestionCount}
                onChange={(event) => setSuggestionCount(Number(event.target.value))}
              />
            </label>
            <button
              type="button"
              className="secondary-button danger-subtle"
              onClick={resetUsage}
            >
              <RotateCcw size={15} />
              {isEn ? "Reset recommendation history" : "Đặt lại lịch sử đề xuất"}
            </button>
          </div>

          <div className="settings-section">
            <div className="settings-section-title">
              <Keyboard size={18} />
              <div>
                <h3>{isEn ? "Formula hotkeys" : "Phím nóng công thức"}</h3>
              </div>
            </div>
            <button
              type="button"
              className="secondary-button settings-hotkey-button"
              onClick={onOpenFormulaHotkeys}
            >
              <Keyboard size={15} />
              {isEn ? "Manage formula hotkeys" : "Quản lý phím nóng công thức"}
            </button>
          </div>

          <div className="settings-section">
            <div className="settings-section-title">
              <SlidersHorizontal size={18} />
              <div>
                <h3>{isEn ? "Appearance & editor" : "Giao diện & biên tập"}</h3>
              </div>
            </div>
            <div className="editor-layout-setting">
              <span>
                <strong>{isEn ? "Editor layout" : "Bố cục soạn thảo"}</strong>
              </span>
              <div
                className="theme-segment editor-layout-segment"
                role="group"
                aria-label={isEn ? "Editor layout" : "Bố cục soạn thảo"}
              >
                <button
                  type="button"
                  className={editorLayout === "standard" ? "is-active" : ""}
                  aria-pressed={editorLayout === "standard"}
                  data-editor-layout-choice="standard"
                  onClick={() => setEditorLayout("standard")}
                >
                  {isEn ? "Standard" : "Tiêu chuẩn"}
                </button>
                <button
                  type="button"
                  className={editorLayout === "classic" ? "is-active" : ""}
                  aria-pressed={editorLayout === "classic"}
                  data-editor-layout-choice="classic"
                  onClick={() => setEditorLayout("classic")}
                >
                  {isEn ? "Classic" : "Cổ điển"}
                </button>
              </div>
            </div>
            <div className="theme-choice-setting">
              <span>
                <strong>{isEn ? "Colour theme" : "Chủ đề màu sắc"}</strong>
              </span>
              <div
                className="theme-segment theme-choice-segment"
                role="group"
                aria-label={isEn ? "Colour theme" : "Chủ đề màu sắc"}
              >
                {THEME_DEFINITIONS.map((definition) => {
                  const swatches =
                    definition.id === "custom"
                      ? ([
                          customTheme.colors.background,
                          customTheme.colors.surface,
                          customTheme.colors.accent,
                        ] as const)
                      : definition.swatches;
                  return (
                    <button
                      key={definition.id}
                      type="button"
                      className={theme === definition.id ? "is-active" : ""}
                      aria-pressed={theme === definition.id}
                      data-theme-choice={definition.id}
                      onClick={() => selectTheme(definition.id)}
                    >
                      <span className="theme-choice-swatch" aria-hidden="true">
                        {swatches.map((color) => (
                          <i key={color} style={{ background: color }} />
                        ))}
                      </span>
                      <span>{isEn ? definition.labelEn : definition.labelVi}</span>
                    </button>
                  );
                })}
              </div>
            </div>
            <label className="switch-row">
              <span>
                <strong>
                  {isEn ? "Auto-pair delimiters" : "Dấu phân cách tự động ghép nối"}
                </strong>
              </span>
              <input
                type="checkbox"
                checked={autoPairDelimiters}
                onChange={(event) =>
                  setAutoPairDelimiters(event.target.checked)
                }
              />
              <span className="switch-control" />
            </label>
            <label className="range-setting">
              <span>
                <strong>{isEn ? "Formula zoom" : "Công thức thu phóng"}</strong>
                <small>{Math.round(zoom * 100)}%</small>
              </span>
              <input
                type="range"
                min="0.5"
                max="1.6"
                step={EDITOR_ZOOM_STEP}
                value={zoom}
                onChange={(event) => setZoom(Number(event.target.value))}
              />
            </label>
          </div>

          <div className="settings-section">
            <button
              type="button"
              className="settings-subdialog-trigger"
              data-interface-customization-trigger
              onClick={openInterfaceCustomization}
            >
              <Eye size={18} />
              <span>
                <strong>{isEn ? "Interface customization" : "Tùy chỉnh giao diện"}</strong>
              </span>
              <ChevronRight size={18} aria-hidden="true" />
            </button>
          </div>

          <div className="settings-section" data-configuration-transfer>
            <div className="settings-section-title">
              <Download size={18} />
              <div>
                <h3>{isEn ? "Configuration backup" : "Sao lưu cấu hình"}</h3>
              </div>
            </div>
            <div className="configuration-transfer-actions">
              <button
                type="button"
                className="secondary-button"
                data-save-configuration
                disabled={configurationBusy}
                onClick={() => void saveConfiguration()}
              >
                <Download size={15} />
                {isEn ? "Save current configuration" : "Lưu cấu hình hiện tại"}
              </button>
              <button
                type="button"
                className="secondary-button"
                data-import-configuration
                disabled={configurationBusy}
                onClick={() => configurationInputRef.current?.click()}
              >
                <Upload size={15} />
                {isEn ? "Import configuration" : "Nhập cấu hình"}
              </button>
              <input
                ref={configurationInputRef}
                className="configuration-file-input"
                type="file"
                accept={`.${VISUALTEX_CONFIGURATION_EXTENSION},application/json`}
                onChange={importConfigurationFile}
                tabIndex={-1}
              />
            </div>
            {configurationStatus && (
              <p className="configuration-transfer-status" role="status">
                {configurationStatus}
              </p>
            )}
          </div>

          <div className="settings-section">
            <div className="settings-section-title">
              <Languages size={18} />
              <div>
                <h3>{isEn ? "Interface language" : "Ngôn ngữ giao diện"}</h3>
              </div>
            </div>
            <div className="theme-segment">
              <button
                type="button"
                className={language === "vi" ? "is-active" : ""}
                aria-pressed={language === "vi"}
                onClick={() => setLanguage("vi")}
              >
                Tiếng Việt
              </button>
              <button
                type="button"
                className={language === "en" ? "is-active" : ""}
                aria-pressed={language === "en"}
                onClick={() => setLanguage("en")}
              >
                English
              </button>
            </div>
          </div>

          <div className="settings-section">
            <div className="settings-section-title">
              <Presentation size={18} />
              <div>
                <h3>
                  {isEn
                    ? "PowerPoint formula defaults"
                    : "Mặc định công thức PowerPoint"}
                </h3>
              </div>
            </div>
            <label className="number-setting office-default-font-size-setting">
              <span>
                <strong>
                  {isEn ? "Default formula font size" : "Cỡ chữ công thức mặc định"}
                </strong>
                <small>
                  {powerPointDefaultFontSizePt} {isEn ? "pt" : "điểm"}
                </small>
              </span>
              <input
                type="number"
                min="5"
                max="200"
                step="0.5"
                value={powerPointDefaultFontSizePt}
                data-powerpoint-default-font-size
                onChange={(event) => {
                  const value = Number(event.target.value);
                  if (Number.isFinite(value)) {
                    setPowerPointDefaultFontSizePt(value);
                  }
                }}
              />
            </label>
            <div
              className="theme-segment office-font-size-presets"
              role="group"
              aria-label={
                isEn
                  ? "PowerPoint formula size presets"
                  : "Cài đặt trước kích thước công thức PowerPoint"
              }
            >
              {[14, 16, 18, 20, 24, 28, 32, 36].map((fontSizePt) => (
                <button
                  key={fontSizePt}
                  type="button"
                  className={
                    powerPointDefaultFontSizePt === fontSizePt
                      ? "is-active"
                      : ""
                  }
                  aria-pressed={powerPointDefaultFontSizePt === fontSizePt}
                  onClick={() => setPowerPointDefaultFontSizePt(fontSizePt)}
                >
                  {fontSizePt}
                </button>
              ))}
            </div>
          </div>

          <OfficeIntegrationSettings />

          <div className="settings-section">
            <div className="settings-section-title">
              <RefreshCw size={18} />
              <div>
                <h3>{isEn ? "Application updates" : "Cập nhật ứng dụng"}</h3>
              </div>
            </div>
            <label className="switch-row">
              <span>
                <strong>{isEn ? "Automatic update notifications" : "Thông báo cập nhật tự động"}</strong>
              </span>
              <input
                type="checkbox"
                checked={checkUpdatesOnStartup}
                onChange={(event) =>
                  setCheckUpdatesOnStartup(event.target.checked)
                }
              />
              <span className="switch-control" />
            </label>
            <button
              type="button"
              className="secondary-button settings-update-button"
              onClick={onCheckForUpdates}
            >
              <RefreshCw size={15} />
              {isEn ? "Check now" : "Kiểm tra ngay"}
            </button>
          </div>
        </div>

        <footer className="dialog-footer">
          <span>{isEn ? "Settings saved automatically" : "Cài đặt được lưu tự động"}</span>
          <button type="button" className="primary-button" onClick={onClose}>
            {isEn ? "Done" : "Xong"}
          </button>
        </footer>

        {interfaceCustomizationOpen &&
          createPortal(
            <div
            className="settings-subdialog-backdrop"
            role="presentation"
            onMouseDown={(event) => {
              event.stopPropagation();
              closeInterfaceCustomization();
            }}
          >
            <section
              ref={interfaceCustomizationDialogRef}
              className="settings-subdialog"
              role="dialog"
              aria-modal="true"
              aria-labelledby="interface-customization-title"
              data-interface-customization-dialog
              onMouseDown={(event) => event.stopPropagation()}
            >
              <header className="dialog-header">
                <div>
                  <span className="eyebrow">INTERFACE</span>
                  <h2 id="interface-customization-title">
                    {isEn ? "Interface customization" : "Tùy chỉnh giao diện"}
                  </h2>
                </div>
                <button
                  type="button"
                  className="icon-button"
                  data-interface-customization-close
                  onClick={closeInterfaceCustomization}
                  aria-label={
                    isEn
                      ? "Close interface customization"
                      : "Đóng tùy chỉnh giao diện"
                  }
                >
                  <X size={18} />
                </button>
              </header>

              <div className="settings-subdialog-content">
                <section
                  className="theme-studio-customization"
                  aria-labelledby="theme-studio-title"
                  data-theme-studio
                >
                  <header className="theme-studio-header">
                    <div>
                      <strong id="theme-studio-title">
                        {isEn ? "Theme studio" : "Studio chủ đề"}
                      </strong>
                    </div>
                    <div className="theme-studio-actions">
                      <button
                        type="button"
                        className="secondary-button"
                        data-theme-reset-preset
                        onClick={resetThemePalette}
                      >
                        {isEn ? "Reset custom" : "Đặt lại tùy chỉnh"}
                      </button>
                      <button
                        type="button"
                        className={theme === "custom" ? "primary-button" : "secondary-button"}
                        data-theme-use-custom
                        onClick={() => selectTheme("custom")}
                      >
                        {isEn ? "Use Custom" : "Sử dụng tùy chỉnh"}
                      </button>
                    </div>
                  </header>

                  <div className="theme-preset-section">
                    <strong>{isEn ? "Ready-made palettes" : "Bảng màu làm sẵn"}</strong>
                    <div className="theme-preset-grid" role="group">
                      {THEME_DEFINITIONS.filter((preset) => preset.id !== "custom").map(
                        (preset) => (
                          <button
                            type="button"
                            key={preset.id}
                            className={theme === preset.id ? "is-active" : ""}
                            aria-pressed={theme === preset.id}
                            data-theme-preset={preset.id}
                            onClick={() => selectTheme(preset.id)}
                          >
                            <span
                              className="theme-preset-aa"
                              style={{
                                background: preset.swatches[0],
                                color: preset.swatches[2],
                                borderColor: preset.swatches[1],
                              }}
                              aria-hidden="true"
                            >
                              Aa
                            </span>
                            <span className="theme-preset-name">
                              {isEn ? preset.labelEn : preset.labelVi}
                            </span>
                            <span className="theme-preset-swatches" aria-hidden="true">
                              {preset.swatches.map((color) => (
                                <i key={color} style={{ background: color }} />
                              ))}
                            </span>
                          </button>
                        ),
                      )}
                    </div>
                  </div>

                  <div className="theme-mode-row">
                    <strong>{isEn ? "Control appearance" : "Ngoại hình điều khiển"}</strong>
                    <div className="theme-segment theme-mode-segment" role="group">
                      <button
                        type="button"
                        className={customTheme.mode === "light" ? "is-active" : ""}
                        aria-pressed={customTheme.mode === "light"}
                        onClick={() =>
                          updateCustomTheme((current) => ({
                            ...current,
                            mode: "light",
                          }))
                        }
                      >
                        {isEn ? "Light" : "Ánh sáng"}
                      </button>
                      <button
                        type="button"
                        className={customTheme.mode === "dark" ? "is-active" : ""}
                        aria-pressed={customTheme.mode === "dark"}
                        onClick={() =>
                          updateCustomTheme((current) => ({
                            ...current,
                            mode: "dark",
                          }))
                        }
                      >
                        {isEn ? "Dark" : "Tối"}
                      </button>
                    </div>
                  </div>

                  <details className="theme-color-details" open>
                    <summary>{isEn ? "Interface colors" : "Màu giao diện"}</summary>
                    {renderThemeColorFields(THEME_CORE_COLOR_FIELDS)}
                  </details>

                  <details className="theme-color-details" open>
                    <summary>{isEn ? "Formula colors" : "Công thức màu sắc"}</summary>
                    {renderThemeColorFields(THEME_FORMULA_COLOR_FIELDS)}
                  </details>

                  <details className="theme-color-details">
                    <summary>{isEn ? "Source editor colors" : "Màu của trình soạn thảo nguồn"}</summary>
                    {renderThemeColorFields(THEME_SYNTAX_COLOR_FIELDS)}
                  </details>

                  <details className="theme-color-details">
                    <summary>{isEn ? "Toolbar category colors" : "Màu sắc danh mục trên Thanh công cụ"}</summary>
                    {renderThemeColorFields(THEME_TOOLBAR_COLOR_FIELDS)}
                  </details>

                  <div
                    className="theme-studio-preview"
                    data-theme-studio-preview
                    style={
                      {
                        "--theme-preview-accent": customTheme.colors.accent,
                        "--theme-preview-accent-hover": customTheme.colors.accentHover,
                        "--theme-preview-accent-soft": customTheme.colors.accentSoft,
                        "--theme-preview-bg": customTheme.colors.background,
                        "--theme-preview-elevated": customTheme.colors.elevated,
                        "--theme-preview-surface": customTheme.colors.surface,
                        "--theme-preview-sunken": customTheme.colors.sunken,
                        "--theme-preview-hover": customTheme.colors.hover,
                        "--theme-preview-active": customTheme.colors.active,
                        "--theme-preview-text": customTheme.colors.foreground,
                        "--theme-preview-muted": customTheme.colors.textMuted,
                        "--theme-preview-faint": customTheme.colors.textFaint,
                        "--theme-preview-border": customTheme.colors.border,
                        "--theme-preview-border-strong": customTheme.colors.borderStrong,
                        "--theme-preview-formula": customTheme.colors.formulaSurface,
                        "--theme-preview-placeholder": customTheme.colors.formulaPlaceholder,
                        "--theme-preview-danger": customTheme.colors.danger,
                        "--theme-preview-warning": customTheme.colors.warning,
                        "--theme-preview-success": customTheme.colors.success,
                      } as CSSProperties
                    }
                  >
                    <div className="theme-preview-titlebar">
                      <span className="theme-preview-window-dots" aria-hidden="true">
                        <i />
                        <i />
                        <i />
                      </span>
                      <strong>{isEn ? "Custom" : "Tùy chỉnh"}</strong>
                      <span className="theme-preview-status">
                        {theme === "custom"
                          ? isEn
                            ? "Active"
                            : "Đang hoạt động"
                          : isEn
                            ? "Preview"
                            : "Xem trước"}
                      </span>
                    </div>
                    <div className="theme-preview-body">
                      <aside className="theme-preview-sidebar" aria-hidden="true">
                        <span className="is-active" />
                        <span />
                        <span />
                        <span />
                      </aside>
                      <div className="theme-preview-main">
                        <div className="theme-preview-toolbar">
                          <span />
                          <span />
                          <span />
                          <button type="button" tabIndex={-1}>Aa</button>
                        </div>
                        <div className="theme-preview-editor">
                          <div className="theme-preview-formula-card">
                            <span className="theme-preview-formula-text">x² + y² = r²</span>
                            <i />
                          </div>
                          <div className="theme-preview-controls">
                            <button type="button" tabIndex={-1}>
                              {isEn ? "Primary" : "Tiểu học"}
                            </button>
                            <span className="success" />
                            <span className="warning" />
                            <span className="danger" />
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                </section>

                <label className="switch-row">
                  <span>
                    <strong>
                      {isEn
                        ? "Highlight formula rows"
                        : "Đánh dấu các hàng công thức"}
                    </strong>
                  </span>
                  <input
                    type="checkbox"
                    checked={highlightActiveLine}
                    data-highlight-active-line-setting
                    onChange={(event) =>
                      setHighlightActiveLine(event.target.checked)
                    }
                  />
                  <span className="switch-control" />
                </label>

                <label className="switch-row">
                  <span>
                    <strong>
                      {isEn
                        ? "Show formula row numbers"
                        : "Hiển thị số hàng công thức"}
                    </strong>
                  </span>
                  <input
                    type="checkbox"
                    checked={showLineNumbers}
                    data-show-line-numbers-setting
                    onChange={(event) =>
                      setShowLineNumbers(event.target.checked)
                    }
                  />
                  <span className="switch-control" />
                </label>

                <label className="switch-row">
                  <span>
                    <strong>
                      {isEn
                        ? "Minimize after keypad copy"
                        : "Thu nhỏ sau khi sao chép bàn phím"}
                    </strong>
                    <small>
                      {isEn
                        ? "In keypad mode, Ctrl+S copies the current LaTeX format and then minimizes VisualTeX."
                        : "Ở chế độ bàn phím, Ctrl+S sao chép định dạng LaTeX hiện tại rồi thu nhỏ VisualTeX."}
                    </small>
                  </span>
                  <input
                    type="checkbox"
                    checked={keypadMinimizeOnCopy}
                    data-keypad-minimize-on-copy-setting
                    onChange={(event) =>
                      setKeypadMinimizeOnCopy(event.target.checked)
                    }
                  />
                  <span className="switch-control" />
                </label>

                <section
                  className="formula-toolbar-customization"
                  aria-labelledby="formula-toolbar-customization-title"
                >
                  <header className="formula-inset-customization-header">
                    <div>
                      <strong id="formula-toolbar-customization-title">
                        {isEn
                          ? "Formula toolbar buttons"
                          : "Các nút trên thanh công cụ công thức"}
                      </strong>
                    </div>
                    <button
                      type="button"
                      className="secondary-button formula-inset-reset"
                      data-formula-tool-button-reset="true"
                      onClick={() => {
                        setFormulaToolButtonSize(
                          DEFAULT_FORMULA_TOOL_BUTTON_SIZE,
                        );
                        setFormulaToolButtonPadding(
                          DEFAULT_FORMULA_TOOL_BUTTON_PADDING,
                        );
                      }}
                    >
                      {isEn ? "Reset" : "Đặt lại"}
                    </button>
                  </header>

                  <div className="formula-inset-range-grid">
                    <label className="range-setting formula-inset-range">
                      <span>
                        <strong>{isEn ? "Button size" : "Kích thước nút"}</strong>
                        <small>{formulaToolButtonSize}px</small>
                      </span>
                      <input
                        type="range"
                        min={MIN_FORMULA_TOOL_BUTTON_SIZE}
                        max={MAX_FORMULA_TOOL_BUTTON_SIZE}
                        step="1"
                        value={formulaToolButtonSize}
                        data-formula-tool-button-size-setting
                        onChange={(event) =>
                          setFormulaToolButtonSize(Number(event.target.value))
                        }
                      />
                    </label>
                    <label className="range-setting formula-inset-range">
                      <span>
                        <strong>
                          {isEn ? "Content inset" : "Chèn nội dung"}
                        </strong>
                        <small>{formulaToolButtonPadding}px</small>
                      </span>
                      <input
                        type="range"
                        min={MIN_FORMULA_TOOL_BUTTON_PADDING}
                        max={MAX_FORMULA_TOOL_BUTTON_PADDING}
                        step="1"
                        value={formulaToolButtonPadding}
                        data-formula-tool-button-padding-setting
                        onChange={(event) =>
                          setFormulaToolButtonPadding(
                            Number(event.target.value),
                          )
                        }
                      />
                    </label>
                  </div>
                </section>

                <section
                  className="formula-inset-customization"
                  aria-labelledby="formula-inset-customization-title"
                >
                  <header className="formula-inset-customization-header">
                    <div>
                      <strong id="formula-inset-customization-title">
                        {isEn
                          ? "Formula area spacing"
                          : "Khoảng cách vùng công thức"}
                      </strong>
                    </div>
                    <button
                      type="button"
                      className="secondary-button formula-inset-reset"
                      data-formula-inset-reset="true"
                      onClick={() => {
                        setFormulaInsetLeft(DEFAULT_FORMULA_INSET);
                        setFormulaInsetRight(DEFAULT_FORMULA_INSET);
                        setFormulaRowVerticalInset(
                          DEFAULT_FORMULA_ROW_VERTICAL_INSET,
                        );
                      }}
                    >
                      {isEn ? "Reset" : "Đặt lại"}
                    </button>
                  </header>

                  <div className="formula-inset-range-grid">
                    <label className="range-setting formula-inset-range">
                      <span>
                        <strong>{isEn ? "Left spacing" : "Khoảng cách trái"}</strong>
                        <small>{formulaInsetLeft}px</small>
                      </span>
                      <input
                        type="range"
                        min={MIN_FORMULA_INSET}
                        max={MAX_FORMULA_INSET}
                        step="1"
                        value={formulaInsetLeft}
                        data-formula-inset-left-setting
                        onChange={(event) =>
                          setFormulaInsetLeft(Number(event.target.value))
                        }
                      />
                    </label>
                    <label className="range-setting formula-inset-range">
                      <span>
                        <strong>{isEn ? "Right spacing" : "Khoảng cách phải"}</strong>
                        <small>{formulaInsetRight}px</small>
                      </span>
                      <input
                        type="range"
                        min={MIN_FORMULA_INSET}
                        max={MAX_FORMULA_INSET}
                        step="1"
                        value={formulaInsetRight}
                        data-formula-inset-right-setting
                        onChange={(event) =>
                          setFormulaInsetRight(Number(event.target.value))
                        }
                      />
                    </label>
                    <label className="range-setting formula-inset-range">
                      <span>
                        <strong>
                          {isEn ? "Formula row vertical spacing" : "Khoảng cách dọc hàng công thức"}
                        </strong>
                        <small>{formulaRowVerticalInset}px</small>
                      </span>
                      <input
                        type="range"
                        min={MIN_FORMULA_ROW_VERTICAL_INSET}
                        max={MAX_FORMULA_ROW_VERTICAL_INSET}
                        step="1"
                        value={formulaRowVerticalInset}
                        data-formula-row-vertical-inset-setting
                        onChange={(event) =>
                          setFormulaRowVerticalInset(Number(event.target.value))
                        }
                      />
                    </label>
                  </div>

                  <div
                    className="formula-inset-preview-window"
                    data-formula-inset-preview
                    style={
                      {
                        "--preview-formula-inset-left": `${formulaInsetLeft}px`,
                        "--preview-formula-inset-right": `${formulaInsetRight}px`,
                        "--preview-toolbar-button-size": `${formulaToolButtonSize}px`,
                        "--preview-toolbar-button-padding": `${formulaToolButtonPadding}px`,
                        "--preview-formula-row-vertical-inset": `${formulaRowVerticalInset}px`,
                      } as CSSProperties
                    }
                  >
                    <div className="formula-inset-preview-titlebar">
                      <i />
                      <i />
                      <i />
                      <span>{isEn ? "Live preview" : "Xem trước trực tiếp"}</span>
                    </div>
                    <div className="formula-inset-preview-toolbar">
                      <span>
                        <MathPreview
                          latex={String.raw`\frac{a}{b}`}
                          fit
                          maximumFitScale={1}
                          fitInsetRatio={0.9}
                        />
                      </span>
                      <span>
                        <MathPreview
                          latex={String.raw`\sqrt{x}`}
                          fit
                          maximumFitScale={1}
                          fitInsetRatio={0.9}
                        />
                      </span>
                      <span>
                        <MathPreview
                          latex={String.raw`\int_a^b`}
                          fit
                          maximumFitScale={1}
                          fitInsetRatio={0.9}
                        />
                      </span>
                    </div>
                    <div className="formula-inset-preview-canvas">
                      <div className="formula-inset-preview-row">
                        <MathPreview
                          latex={String.raw`x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}`}
                          fit
                          maximumFitScale={1}
                          fitInsetRatio={0.88}
                        />
                      </div>
                    </div>
                  </div>
                </section>

                <section
                  className="formula-font-customization"
                  aria-labelledby="formula-font-customization-title"
                >
                  <header className="formula-inset-customization-header">
                    <div>
                      <strong id="formula-font-customization-title">
                        {isEn ? "Formula fonts" : "Phông chữ công thức"}
                      </strong>
                    </div>
                    <button
                      type="button"
                      className="secondary-button formula-inset-reset"
                      data-formula-font-reset
                      onClick={() => {
                        setFormulaLetterFont(DEFAULT_FORMULA_LETTER_FONT);
                        setFormulaChineseFont(DEFAULT_FORMULA_CHINESE_FONT);
                      }}
                    >
                      {isEn ? "Reset" : "Đặt lại"}
                    </button>
                  </header>

                  <div className="formula-font-select-grid">
                    <label className="formula-font-setting">
                      <span>
                        <strong>{isEn ? "Math letter font" : "Phông chữ toán học"}</strong>
                      </span>
                      <select
                        value={formulaLetterFont}
                        data-formula-letter-font-setting
                        onChange={(event) =>
                          setFormulaLetterFont(
                            event.currentTarget.value as FormulaLetterFont,
                          )
                        }
                      >
                        {FORMULA_LETTER_FONT_OPTIONS.map((item) => (
                          <option key={item.id} value={item.id}>
                            {item.label}
                          </option>
                        ))}
                      </select>
                    </label>
                  </div>

                  <div
                    className="formula-font-live-preview"
                    data-formula-font-preview
                    style={
                      {
                        "--visualtex-formula-upright-font-family":
                          formulaLetterFamilies.upright,
                        "--visualtex-formula-italic-font-family":
                          formulaLetterFamilies.italic,
                      } as CSSProperties
                    }
                  >
                    <span>{isEn ? "Live preview" : "Xem trước trực tiếp"}</span>
                    <MathPreview
                      latex={String.raw`E=mc^2,\quad \alpha+\beta=\gamma`}
                      staticLayout
                    />
                  </div>
                </section>

                <section
                  className="png-background-customization"
                  aria-labelledby="png-background-customization-title"
                >
                  <header className="formula-inset-customization-header">
                    <div>
                      <strong id="png-background-customization-title">
                        {isEn ? "PNG background" : "PNG nền"}
                      </strong>
                    </div>
                  </header>
                  <div className="png-background-controls">
                    <button
                      type="button"
                      className={`png-background-transparent${pngExportBackground === "transparent" ? " is-active" : ""}`}
                      data-png-background-transparent
                      aria-pressed={pngExportBackground === "transparent"}
                      onClick={() => setPngExportBackground("transparent")}
                    >
                      <span className="png-transparent-swatch" aria-hidden="true" />
                      <span>
                        <strong>{isEn ? "Transparent" : "Trong suốt"}</strong>
                      </span>
                    </button>
                    <label
                      className={`png-background-color${pngExportBackground !== "transparent" ? " is-active" : ""}`}
                    >
                      <input
                        type="color"
                        value={pngExportBackgroundPickerValue(pngExportBackground)}
                        data-png-background-color-setting
                        onChange={(event) =>
                          setPngExportBackground(event.currentTarget.value as `#${string}`)
                        }
                      />
                      <span>
                        <strong>{isEn ? "Custom colour" : "Màu tùy chỉnh"}</strong>
                      </span>
                    </label>
                  </div>
                </section>
              </div>

              <footer className="dialog-footer">
                <span>
                  {isEn
                    ? "Changes apply immediately"
                    : "Thay đổi được áp dụng ngay lập tức"}
                </span>
                <button
                  type="button"
                  className="primary-button"
                  onClick={closeInterfaceCustomization}
                >
                  {isEn ? "Done" : "Xong"}
                </button>
              </footer>
            </section>
          </div>,
            document.body,
          )}
      </section>
    </div>
  );
}
