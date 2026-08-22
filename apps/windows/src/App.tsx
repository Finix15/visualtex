import { ChangeEvent, useCallback, useEffect, useRef, useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import {
  AlertCircle,
  BookOpenText,
  Braces,
  Check,
  ChevronDown,
  Code2,
  FileDown,
  FilePlus2,
  FolderOpen,
  History,
  Keyboard,
  Languages,
  LoaderCircle,
  Menu,
  Minus,
  PanelBottomClose,
  PanelBottomOpen,
  PanelLeftClose,
  PanelLeftOpen,
  PanelRightClose,
  PanelRightOpen,
  Plus,
  Redo2,
  RefreshCw,
  Save,
  ScanLine,
  Settings2,
  Undo2,
  X,
} from "lucide-react";
import {
  MathEditor,
  type MathEditorHandle,
  type MathEditorInsertionTarget,
} from "./editor/MathEditor";
import { FormulaToolbar } from "./toolbar/FormulaToolbar";
import { LatexSourceEditor } from "./source-editor/LatexSourceEditor";
import { SettingsDialog } from "./components/SettingsDialog";
import { FormulaHotkeyManagerDialog } from "./components/FormulaHotkeyManagerDialog";
import { HistoryPanel } from "./components/HistoryPanel";
import { HelpDialog } from "./components/HelpDialog";
import { OcrDialog } from "./components/OcrDialog";
import { ExportDialog } from "./components/ExportDialog";
import { UpdateDialog } from "./components/UpdateDialog";
import { VisualTeXLogo } from "./components/VisualTeXLogo";
import { EditorWorkspace } from "./workspace/EditorWorkspace";
import {
  EDITOR_ZOOM_STEP,
  MAX_EDITOR_ZOOM,
  MIN_EDITOR_ZOOM,
  joinFormulaLines,
  useEditorStore,
} from "./stores/editorStore";
import {
  historyManager,
  useHistorySnapshot,
} from "./history/HistoryManager";
import {
  applyHistoryEntryToEditor,
  createBlankDocumentSnapshot,
  documentSnapshotsEquivalent,
  getEditorDocumentSnapshot,
  reconcileFormulaLines,
} from "./history/documentHistory";
import type {
  DocumentSnapshot,
  ReplaceDocumentEntry,
} from "./history/historyTypes";
import {
  copyLatex,
  formatLatex,
  getLatexCodeFormatDefinition,
  latexCodeFormats,
  parseLatexSource,
} from "./clipboard/LatexCopyService";
import { normalizeChineseLatex } from "./editor/normalizeChineseLatex";
import type { FormulaDocument, LatexCodeFormat } from "./types/formula";
import { applyDocumentTheme, publishSynchronizedTheme } from "./themeSync";
import {
  CUSTOM_THEME_CHANGED_EVENT,
  readCustomTheme,
} from "./themeCustomization";
import { copyFormulaDocumentPngToClipboard } from "./export/pngClipboard";
import {
  DEFAULT_OCR_MODEL,
  OCR_MODELS,
  cancelOcrRecognition,
  fileToOcrRequest,
  getOcrRuntimeStatus,
  isTauriEnvironment,
  listenOcrRecognitionProgress,
  recognizeFormulaImage,
  warmupOcrModel,
  type OcrModelName,
} from "./ocr/ocrService";
import {
  checkForUpdates,
  openReleasePage,
  PROJECT_URL,
  type UpdateCheckResult,
} from "./update/updateService";
import { readLocalStorage, writeLocalStorage } from "./runtime/safeStorage";
import {
  captureQuickOcrScreenshot,
  configureSilentOcr,
  isQuickOcrCaptureMode,
  quickOcrCaptureToFile,
  restoreQuickOcrWindow,
  QUICK_OCR_CAPTURE_MODE_STORAGE_KEY,
  SILENT_OCR_SHORTCUT,
  SILENT_OCR_STORAGE_KEY,
  waitForQuickOcrSystemScreenshot,
  type QuickOcrCaptureMode,
} from "./ocr/quickOcr";

type InlineOcrStatus = "running" | "cancelling" | "success" | "error" | "cancelled";

interface InlineOcrState {
  status: InlineOcrStatus;
  message: string;
  seconds: number;
  model: OcrModelName;
}

const OCR_MODEL_STORAGE_KEY = "visualtex.ocr.model";

function App() {
  const editorRef = useRef<MathEditorHandle>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const menuButtonRef = useRef<HTMLButtonElement>(null);
  const appMenuRef = useRef<HTMLDivElement>(null);
  const copyMenuButtonRef = useRef<HTMLButtonElement>(null);
  const copyMenuRef = useRef<HTMLDivElement>(null);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [formulaHotkeyManagerOpen, setFormulaHotkeyManagerOpen] = useState(false);
  const [historyOpen, setHistoryOpen] = useState(false);
  const [helpOpen, setHelpOpen] = useState(false);
  const [ocrOpen, setOcrOpen] = useState(false);
  const [exportOpen, setExportOpen] = useState(false);
  const [sidebarOpen, setSidebarOpen] = useState(() => window.innerWidth >= 1040);
  const [copyMenuOpen, setCopyMenuOpen] = useState(false);
  const [menuOpen, setMenuOpen] = useState(false);
  const [keypadMode, setKeypadMode] = useState(false);
  const [updateOpen, setUpdateOpen] = useState(false);
  const [updateChecking, setUpdateChecking] = useState(false);
  const [updateError, setUpdateError] = useState("");
  const [updateResult, setUpdateResult] = useState<UpdateCheckResult | null>(null);
  const [automaticUpdatePrompt, setAutomaticUpdatePrompt] = useState(false);
  const [toast, setToast] = useState("");
  const [savedPulse, setSavedPulse] = useState(false);
  const [editorHistoryBusy, setEditorHistoryBusy] = useState(false);
  const [ocrModel, setOcrModel] = useState<OcrModelName>(() => {
    const stored = readLocalStorage(OCR_MODEL_STORAGE_KEY);
    return OCR_MODELS.some((item) => item.id === stored)
      ? (stored as OcrModelName)
      : DEFAULT_OCR_MODEL;
  });
  const [silentOcrEnabled, setSilentOcrEnabled] = useState(
    () => readLocalStorage(SILENT_OCR_STORAGE_KEY) === "true",
  );
  const [silentOcrShortcut, setSilentOcrShortcut] = useState(SILENT_OCR_SHORTCUT);
  const [quickOcrCaptureMode, setQuickOcrCaptureMode] = useState<QuickOcrCaptureMode>(() => {
    const stored = readLocalStorage(QUICK_OCR_CAPTURE_MODE_STORAGE_KEY);
    return isQuickOcrCaptureMode(stored) ? stored : "immediate";
  });
  const [quickOcrCaptureBusy, setQuickOcrCaptureBusy] = useState(false);
  const [inlineOcr, setInlineOcr] = useState<InlineOcrState | null>(null);
  const startupOcrModelRef = useRef(ocrModel);
  const inlineOcrBusyRef = useRef(false);
  const inlineOcrCancelRequestedRef = useRef(false);
  const inlineOcrRunIdRef = useRef(0);
  const inlineOcrClearTimerRef = useRef<number | null>(null);
  const automaticUpdateCheckRef = useRef(false);
  const initialEditorFocusDoneRef = useRef(false);
  const pngClipboardBusyRef = useRef(false);

  const title = useEditorStore((state) => state.title);
  const setTitle = useEditorStore((state) => state.setTitle);
  const lines = useEditorStore((state) => state.lines);
  const activeLineId = useEditorStore((state) => state.activeLineId);
  const formulaAlignment = useEditorStore((state) => state.formulaAlignment);
  const theme = useEditorStore((state) => state.theme);
  const synchronizedThemeRef = useRef(theme);
  const language = useEditorStore((state) => state.language);
  const setLanguage = useEditorStore((state) => state.setLanguage);
  const zoom = useEditorStore((state) => state.zoom);
  const setZoom = useEditorStore((state) => state.setZoom);
  const editorLayout = useEditorStore((state) => state.editorLayout);
  const pngExportBackground = useEditorStore(
    (state) => state.pngExportBackground,
  );
  const formulaLetterFont = useEditorStore((state) => state.formulaLetterFont);
  const formulaChineseFont = useEditorStore((state) => state.formulaChineseFont);
  const latexCodeFormat = useEditorStore((state) => state.latexCodeFormat);
  const setLatexCodeFormat = useEditorStore(
    (state) => state.setLatexCodeFormat,
  );
  const addHistory = useEditorStore((state) => state.addHistory);
  const keypadMinimizeOnCopy = useEditorStore((state) => state.keypadMinimizeOnCopy);
  const loadDocument = useEditorStore((state) => state.loadDocument);
  const toDocument = useEditorStore((state) => state.toDocument);
  const checkUpdatesOnStartup = useEditorStore(
    (state) => state.checkUpdatesOnStartup,
  );
  const setCheckUpdatesOnStartup = useEditorStore(
    (state) => state.setCheckUpdatesOnStartup,
  );
  const powerPointDefaultFontSizePt = useEditorStore(
    (state) => state.powerPointDefaultFontSizePt,
  );
  const historyState = useHistorySnapshot();
  const isEn = language === "en";

  useEffect(() => {
    void invoke("write_office_ui_language", { language }).catch((error) => {
      console.warn("Unable to synchronize the Office UI language", error);
    });
  }, [language]);
  const latex = joinFormulaLines(lines);
  const sourceLatex = formatLatex(latex, latexCodeFormat);
  const currentCodeFormat = getLatexCodeFormatDefinition(latexCodeFormat);
  const codeFormatGroups = [
    {
      id: "single" as const,
      title: isEn ? "Independent formula formats" : "Định dạng công thức độc lập",
      description: isEn
        ? "Each non-empty formula field gets its own wrapper"
        : "Mỗi trường công thức không trống sẽ có trình bao bọc riêng",
      formats: latexCodeFormats.filter((format) => format.group === "single"),
    },
    {
      id: "multi" as const,
      title: isEn ? "Combined multi-line environments" : "Môi trường đa dòng kết hợp",
      description: isEn
        ? "All non-empty formula fields become rows in one environment"
        : "Tất cả các trường công thức không trống sẽ trở thành hàng trong một môi trường",
      formats: latexCodeFormats.filter((format) => format.group === "multi"),
    },
  ];
  const selectedOcrModel =
    OCR_MODELS.find((item) => item.id === ocrModel) ??
    OCR_MODELS.find((item) => item.id === DEFAULT_OCR_MODEL)!;
  const inlineOcrModel =
    OCR_MODELS.find((item) => item.id === inlineOcr?.model) ?? selectedOcrModel;
  const inlineOcrIsBusy =
    inlineOcr?.status === "running" || inlineOcr?.status === "cancelling";

  const captureDocumentSnapshot = (): DocumentSnapshot =>
    getEditorDocumentSnapshot(editorRef.current?.getSelectionMap() ?? {});

  const restoreSnapshotFocus = (snapshot: DocumentSnapshot) => {
    const lineId = snapshot.activeLineId;
    if (!lineId) return;
    const line = snapshot.lines.find((item) => item.id === lineId);
    if (!line) return;
    void editorRef.current?.restoreSelection(
      lineId,
      line.latex,
      snapshot.selectionByLineId[lineId] ?? null,
    );
  };

  const replaceDocumentWithHistory = (
    after: DocumentSnapshot,
    source: ReplaceDocumentEntry["source"],
  ) => {
    if (source !== "source-apply") historyManager.commitPendingTransaction();
    const before = captureDocumentSnapshot();
    if (documentSnapshotsEquivalent(before, after)) return false;
    useEditorStore.getState().replaceDocumentState(after);
    const entry: ReplaceDocumentEntry = {
      type: "replace-document",
      before,
      after,
      source,
      timestamp: Date.now(),
    };
    if (source === "source-apply") {
      historyManager.recordSourceDocumentEdit(entry);
    } else {
      historyManager.push(entry);
      window.requestAnimationFrame(() => restoreSnapshotFocus(after));
    }
    return true;
  };

  useEffect(() => {
    if (initialEditorFocusDoneRef.current) return;
    initialEditorFocusDoneRef.current = true;
    let repairTimer = 0;
    const frame = window.requestAnimationFrame(() => {
      editorRef.current?.focus({ target: "last", moveToEnd: true });
      repairTimer = window.setTimeout(() => {
        editorRef.current?.focus({ target: "last", moveToEnd: true });
      }, 80);
    });
    return () => {
      window.cancelAnimationFrame(frame);
      window.clearTimeout(repairTimer);
    };
  }, []);

  useEffect(() => {
    historyManager.configure({
      getDocumentSnapshot: () =>
        getEditorDocumentSnapshot(editorRef.current?.getSelectionMap() ?? {}),
      applyEntry: async (entry, direction) => {
        const target = applyHistoryEntryToEditor(entry, direction);
        if (!target) return;
        // Yield once so React can mount any line restored by the history entry.
        // Do not wait on requestAnimationFrame here: background desktop windows
        // and headless release checks can throttle animation frames indefinitely,
        // leaving history replay active and the Redo action disabled.
        await new Promise<void>((resolve) => window.setTimeout(resolve, 0));
        await editorRef.current?.restoreSelection(
          target.lineId,
          target.latex,
          target.selection,
        );
      },
    });
    return () => historyManager.configure(null);
  }, []);

  useEffect(() => {
    const checkpointTimer = window.setInterval(() => {
      historyManager.commitPendingTransaction();
      void historyManager.createCheckpoint("autosave");
    }, 30_000);
    const handleBeforeUnload = () => {
      historyManager.commitPendingTransaction();
      void historyManager.createCheckpoint("before-unload");
    };
    window.addEventListener("beforeunload", handleBeforeUnload);
    return () => {
      window.clearInterval(checkpointTimer);
      window.removeEventListener("beforeunload", handleBeforeUnload);
    };
  }, []);

  useEffect(() => {
    if (synchronizedThemeRef.current === theme) {
      applyDocumentTheme(theme);
    } else {
      synchronizedThemeRef.current = theme;
      publishSynchronizedTheme(theme);
    }
    if (isTauriEnvironment()) {
      void invoke<string>("set_app_theme", { theme }).catch(() => undefined);
    }
  }, [theme]);

  useEffect(() => {
    if (!isTauriEnvironment()) return;
    void invoke<string>("set_app_editor_layout", { editorLayout }).catch(
      () => undefined,
    );
  }, [editorLayout]);

  useEffect(() => {
    if (!isTauriEnvironment()) return;
    let timer = 0;
    let disposed = false;
    const publishPreferences = () => {
      window.clearTimeout(timer);
      timer = window.setTimeout(() => {
        if (disposed) return;
        const settings = useEditorStore.getState().toDocument().settings;
        void invoke("set_app_editor_preferences", {
          preferences: {
            settings,
            customTheme: readCustomTheme(),
          },
        }).catch(() => undefined);
      }, 24);
    };
    publishPreferences();
    const unsubscribeStore = useEditorStore.subscribe(publishPreferences);
    window.addEventListener(CUSTOM_THEME_CHANGED_EVENT, publishPreferences);
    return () => {
      disposed = true;
      window.clearTimeout(timer);
      unsubscribeStore();
      window.removeEventListener(CUSTOM_THEME_CHANGED_EVENT, publishPreferences);
    };
  }, []);

  useEffect(() => {
    if (!isTauriEnvironment()) return;
    void invoke<number>("set_powerpoint_default_font_size", {
      fontSizePt: powerPointDefaultFontSizePt,
    }).catch(() => undefined);
  }, [powerPointDefaultFontSizePt]);

  useEffect(() => {
    document.documentElement.lang = isEn ? "en" : "vi-VN";
  }, [isEn]);

  useEffect(() => {
    const compactWindow = window.matchMedia("(max-width: 1040px)");
    const handleCompactWindow = (event: MediaQueryListEvent) => {
      if (event.matches) setSidebarOpen(false);
    };
    compactWindow.addEventListener("change", handleCompactWindow);
    return () => compactWindow.removeEventListener("change", handleCompactWindow);
  }, []);

  useEffect(() => {
    if (!latex.trim()) return;
    const timeout = window.setTimeout(() => addHistory(latex), 1800);
    return () => window.clearTimeout(timeout);
  }, [latex, addHistory]);

  useEffect(() => {
    if (!toast) return;
    const timeout = window.setTimeout(() => setToast(""), 1800);
    return () => window.clearTimeout(timeout);
  }, [toast]);

  useEffect(() => {
    if (!isTauriEnvironment()) return;
    const timer = window.setTimeout(() => {
      void warmupOcrModel(startupOcrModelRef.current).catch(() => undefined);
    }, 50);
    return () => window.clearTimeout(timer);
  }, []);

  useEffect(() => {
    if (!isTauriEnvironment()) return;
    let cancelled = false;
    void configureSilentOcr(silentOcrEnabled, ocrModel, latexCodeFormat)
      .then((shortcut) => {
        if (cancelled || !silentOcrEnabled || !shortcut) return;
        setSilentOcrShortcut(shortcut);
      })
      .catch((error) => {
        if (cancelled || !silentOcrEnabled) return;
        setSilentOcrEnabled(false);
        writeLocalStorage(SILENT_OCR_STORAGE_KEY, "false");
        setToast(
          error instanceof Error
            ? error.message
            : isEn
              ? "Unable to enable silent OCR"
              : "Không thể bật OCR im lặng",
        );
      });
    return () => {
      cancelled = true;
    };
  }, [silentOcrEnabled, ocrModel, latexCodeFormat, isEn]);

  useEffect(() => {
    const menu = menuOpen ? appMenuRef.current : copyMenuOpen ? copyMenuRef.current : null;
    const trigger = menuOpen ? menuButtonRef.current : copyMenuButtonRef.current;
    if (!menu || !trigger) return;

    const items = Array.from(
      menu.querySelectorAll<HTMLButtonElement>('button:not(:disabled)'),
    );
    const frame = window.requestAnimationFrame(() => items[0]?.focus());

    const handleMenuKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.preventDefault();
        setMenuOpen(false);
        setCopyMenuOpen(false);
        trigger.focus({ preventScroll: true });
        return;
      }

      if (event.key !== "ArrowDown" && event.key !== "ArrowUp") return;
      event.preventDefault();
      const currentIndex = items.indexOf(document.activeElement as HTMLButtonElement);
      const direction = event.key === "ArrowDown" ? 1 : -1;
      const nextIndex = currentIndex < 0
        ? 0
        : (currentIndex + direction + items.length) % items.length;
      items[nextIndex]?.focus();
    };

    menu.addEventListener("keydown", handleMenuKeyDown);
    return () => {
      window.cancelAnimationFrame(frame);
      menu.removeEventListener("keydown", handleMenuKeyDown);
    };
  }, [menuOpen, copyMenuOpen]);

  const inlineOcrStatus = inlineOcr?.status;
  useEffect(() => {
    if (inlineOcrStatus !== "running" && inlineOcrStatus !== "cancelling") {
      return;
    }
    const timer = window.setInterval(() => {
      setInlineOcr((current) =>
        current
          ? {
              ...current,
              seconds: current.seconds + 1,
            }
          : current,
      );
    }, 1000);
    return () => window.clearInterval(timer);
  }, [inlineOcrStatus]);

  useEffect(
    () => () => {
      if (inlineOcrClearTimerRef.current !== null) {
        window.clearTimeout(inlineOcrClearTimerRef.current);
      }
    },
    [],
  );

  const scheduleInlineOcrClear = (delay: number) => {
    if (inlineOcrClearTimerRef.current !== null) {
      window.clearTimeout(inlineOcrClearTimerRef.current);
    }
    inlineOcrClearTimerRef.current = window.setTimeout(() => {
      setInlineOcr(null);
      inlineOcrClearTimerRef.current = null;
    }, delay);
  };

  const handleOcrModelChange = (nextModel: OcrModelName) => {
    if (inlineOcrBusyRef.current || nextModel === ocrModel) return;
    startupOcrModelRef.current = nextModel;
    setOcrModel(nextModel);
    writeLocalStorage(OCR_MODEL_STORAGE_KEY, nextModel);
    void warmupOcrModel(nextModel).catch(() => undefined);
  };

  const cancelInlineOcr = async () => {
    if (!inlineOcrBusyRef.current) return;
    inlineOcrCancelRequestedRef.current = true;
    setInlineOcr((current) =>
      current
        ? {
            ...current,
            status: "cancelling",
            message: isEn ? "Cancelling OCR…" : "Đang hủy OCR…",
          }
        : current,
    );
    try {
      await cancelOcrRecognition();
    } catch {
      // The recognition promise will surface the final state. A worker that
      // already exited is equivalent to a successful cancellation.
    }
  };

  const handleEditorImagePaste = async (
    file: File,
    target: MathEditorInsertionTarget,
    source: "paste" | "quick" = "paste",
  ) => {
    if (inlineOcrBusyRef.current) {
      setToast(isEn ? "Another pasted image is being recognized" : "Một hình ảnh được dán khác đang được nhận dạng");
      return;
    }
    if (!isTauriEnvironment()) {
      setToast(isEn ? "Image OCR is available in the desktop app" : "OCR hình ảnh có sẵn trong ứng dụng máy tính để bàn");
      return;
    }

    if (inlineOcrClearTimerRef.current !== null) {
      window.clearTimeout(inlineOcrClearTimerRef.current);
      inlineOcrClearTimerRef.current = null;
    }

    const runId = ++inlineOcrRunIdRef.current;
    inlineOcrBusyRef.current = true;
    inlineOcrCancelRequestedRef.current = false;
    setInlineOcr({
      status: "running",
      message: isEn ? "Checking the local OCR runtime…" : "Đang kiểm tra thời gian chạy OCR cục bộ…",
      seconds: 0,
      model: ocrModel,
    });

    let unlisten: (() => void) | undefined;
    try {
      const runtime = await getOcrRuntimeStatus();
      if (inlineOcrCancelRequestedRef.current) throw new Error("OCR_CANCELLED");
      if (!runtime.installed) {
        setOcrOpen(true);
        throw new Error(
          isEn
            ? "Install the OCR runtime before pasting an image"
            : "Cài đặt thời gian chạy OCR trước khi dán hình ảnh",
        );
      }

      if (!runtime.installedModels.includes(ocrModel)) {
        setOcrOpen(true);
        throw new Error(
          isEn
            ? `Install ${selectedOcrModel.labelEn} before using it for OCR`
            : `Cài đặt ${selectedOcrModel.labelEn} trước khi sử dụng cho OCR`,
        );
      }
      const availableOcrModel = ocrModel;

      unlisten = await listenOcrRecognitionProgress((progress) => {
        if (
          inlineOcrRunIdRef.current !== runId ||
          progress.model !== ocrModel
        ) {
          return;
        }
        setInlineOcr((current) =>
          current
            ? {
                ...current,
                message: progress.message,
              }
            : current,
        );
      });

      const request = await fileToOcrRequest(file, availableOcrModel);
      if (inlineOcrCancelRequestedRef.current) throw new Error("OCR_CANCELLED");
      const result = await recognizeFormulaImage(request);
      if (
        inlineOcrCancelRequestedRef.current ||
        inlineOcrRunIdRef.current !== runId
      ) {
        throw new Error("OCR_CANCELLED");
      }

      const recognizedLatex = result.formulas
        .map((formula) => formula.latex.trim())
        .filter(Boolean)
        .join("\n");
      if (!recognizedLatex) {
        throw new Error(isEn ? "OCR returned an empty formula" : "OCR trả về công thức trống");
      }

      const inserted =
        editorRef.current?.insertLatexAt(target, recognizedLatex, "ocr") ?? false;
      if (!inserted) {
        throw new Error(
          isEn
            ? "The original formula line no longer exists; the OCR result was not inserted"
            : "Dòng công thức gốc không còn tồn tại; kết quả OCR không được chèn",
        );
      }

      setInlineOcr((current) => ({
        status: "success",
        message: result.backgroundInverted
          ? isEn
            ? "Recognized and inserted · dark background inverted"
            : "Đã nhận dạng và chèn · đảo ngược nền tối"
          : isEn
            ? "Recognized and inserted at the saved cursor"
            : "Nhận dạng và chèn vào con trỏ đã lưu",
        seconds: current?.seconds ?? 0,
        model: ocrModel,
      }));
      setToast(
        source === "quick"
          ? isEn
            ? "Screenshot converted to LaTeX"
            : "Ảnh chụp màn hình được chuyển đổi sang LaTeX"
          : isEn
            ? "Pasted image converted to LaTeX"
            : "Hình ảnh đã dán được chuyển đổi sang LaTeX",
      );
      scheduleInlineOcrClear(1800);
    } catch (error) {
      const errorMessage =
        error instanceof Error ? error.message : typeof error === "string" ? error : "";
      const cancelled =
        inlineOcrCancelRequestedRef.current || errorMessage.includes("OCR_CANCELLED");
      if (cancelled) {
        setInlineOcr((current) => ({
          status: "cancelled",
          message: isEn ? "OCR cancelled" : "OCR đã bị hủy",
          seconds: current?.seconds ?? 0,
          model: ocrModel,
        }));
        scheduleInlineOcrClear(1200);
      } else {
        const message =
          errorMessage || (isEn ? "Image OCR failed" : "OCR hình ảnh không thành công");
        setInlineOcr((current) => ({
          status: "error",
          message,
          seconds: current?.seconds ?? 0,
          model: ocrModel,
        }));
        setToast(message);
        scheduleInlineOcrClear(4500);
      }
    } finally {
      unlisten?.();
      if (inlineOcrRunIdRef.current === runId) {
        inlineOcrBusyRef.current = false;
        inlineOcrCancelRequestedRef.current = false;
      }
    }
  };

  const handleQuickOcr = async () => {
    if (inlineOcrBusyRef.current || quickOcrCaptureBusy) {
      setToast(isEn ? "OCR is already running" : "OCR đang chạy");
      return;
    }
    const target = editorRef.current?.captureInsertionTarget();
    if (!target) {
      setToast(
        isEn
          ? "Click a formula field before quick OCR"
          : "Nhấp vào trường công thức trước OCR nhanh",
      );
      return;
    }

    setQuickOcrCaptureBusy(true);
    try {
      const waitingForSystemScreenshot =
        quickOcrCaptureMode === "system-screenshot";
      setToast(
        waitingForSystemScreenshot
          ? isEn
            ? "VisualTeX minimized. Switch to the target page and take a Windows screenshot within 60 seconds."
            : "VisualTeX được thu nhỏ. Chuyển sang trang đích và chụp ảnh màn hình Windows trong vòng 60 giây."
          : isEn
            ? "Select a formula area in the Windows Snipping Tool"
            : "Chọn vùng công thức trong Windows Snipping Tool",
      );
      let capture: Awaited<ReturnType<typeof captureQuickOcrScreenshot>> = null;
      try {
        capture = waitingForSystemScreenshot
          ? await waitForQuickOcrSystemScreenshot()
          : await captureQuickOcrScreenshot();
      } finally {
        // Quick OCR is an interactive main-app workflow: after the screenshot
        // is committed, return to the editor before recognition/insertion so
        // the user sees the progress and the recognized formula immediately.
        await restoreQuickOcrWindow();
      }
      if (!capture) {
        setToast(
          waitingForSystemScreenshot
            ? isEn
              ? "No new Windows screenshot was detected within 60 seconds"
              : "Không phát hiện thấy ảnh chụp màn hình Windows mới trong vòng 60 giây"
            : isEn
              ? "Screenshot cancelled"
              : "Ảnh chụp màn hình bị hủy",
        );
        return;
      }
      await handleEditorImagePaste(
        quickOcrCaptureToFile(capture),
        target,
        "quick",
      );
    } catch (error) {
      setToast(
        error instanceof Error
          ? error.message
          : isEn
            ? "Quick OCR failed"
            : "OCR nhanh không thành công",
      );
    } finally {
      setQuickOcrCaptureBusy(false);
    }
  };

  const handleQuickOcrCaptureModeChange = (mode: QuickOcrCaptureMode) => {
    setQuickOcrCaptureMode(mode);
    writeLocalStorage(QUICK_OCR_CAPTURE_MODE_STORAGE_KEY, mode);
    setToast(
      mode === "system-screenshot"
        ? isEn
          ? "Quick OCR mode: wait for Windows screenshot"
          : "Chế độ OCR nhanh: chờ ảnh chụp màn hình Windows"
        : isEn
          ? "Quick OCR mode: immediate selection"
          : "Chế độ OCR nhanh: chọn ngay lập tức",
    );
  };

  const handleSilentOcrEnabledChange = (enabled: boolean) => {
    setSilentOcrEnabled(enabled);
    writeLocalStorage(SILENT_OCR_STORAGE_KEY, String(enabled));
    setToast(
      enabled
        ? isEn
          ? `Silent OCR enabled · ${silentOcrShortcut}`
          : `Đã bật OCR im lặng · ${silentOcrShortcut}`
        : isEn
          ? "Silent OCR disabled"
          : "OCR im lặng bị tắt",
    );
  };

  useEffect(() => {
    const runQuick = () => void handleQuickOcr();
    window.addEventListener("visualtex-quick-ocr", runQuick);
    return () => {
      window.removeEventListener("visualtex-quick-ocr", runQuick);
    };
  });

  const handleCodeFormatChange = (format: LatexCodeFormat) => {
    const definition = getLatexCodeFormatDefinition(format);
    setLatexCodeFormat(format);
    setCopyMenuOpen(false);
    setToast(
      isEn
        ? `LaTeX code format: ${definition.titleEn}`
        : `Định dạng mã LaTeX: ${definition.titleEn}`,
    );
  };

  const handleCopy = async () => {
    try {
      await copyLatex(latex, latexCodeFormat);
      addHistory(latex);
      setToast(
        isEn
          ? `Copied ${currentCodeFormat.titleEn}`
          : `Sao chép ${currentCodeFormat.titleEn}`,
      );
      return true;
    } catch {
      setToast(
        isEn
          ? "Copy failed. Check clipboard permission."
          : "Sao chép không thành công. Kiểm tra quyền của clipboard.",
      );
      return false;
    }
  };

  const handleKeypadCopy = async () => {
    const copied = await handleCopy();
    if (!copied || !keypadMinimizeOnCopy || !isTauriEnvironment()) return;
    try {
      await invoke("minimize_visualtex_main_window");
    } catch {
      setToast(
        isEn
          ? "LaTeX copied, but VisualTeX could not be minimized."
          : "LaTeX đã sao chép nhưng không thể thu nhỏ VisualTeX.",
      );
    }
  };

  const handleKeypadModeToggle = async () => {
    const nextMode = !keypadMode;
    setMenuOpen(false);
    setCopyMenuOpen(false);
    if (isTauriEnvironment()) {
      try {
        await invoke("set_main_window_keypad_mode", { enabled: nextMode });
      } catch (cause) {
        const message = cause instanceof Error ? cause.message : String(cause);
        setToast(
          isEn
            ? `Unable to switch keypad window size: ${message}`
            : `Không thể chuyển đổi kích thước cửa sổ bàn phím: ${message}`,
        );
        return;
      }
    }
    setKeypadMode(nextMode);
    window.requestAnimationFrame(() =>
      editorRef.current?.focus({ target: "last", moveToEnd: false }),
    );
  };

  const handleCopyPng = async () => {
    if (pngClipboardBusyRef.current) return;
    pngClipboardBusyRef.current = true;
    try {
      await copyFormulaDocumentPngToClipboard(
        lines.map((line) => line.latex),
        {
          background: pngExportBackground,
          formulaLetterFont,
          formulaChineseFont,
        },
      );
      setToast(isEn ? "PNG copied to Clipboard" : "PNG được sao chép vào Clipboard");
    } catch (cause) {
      const message = cause instanceof Error ? cause.message : String(cause);
      setToast(
        isEn
          ? `Unable to copy PNG: ${message}`
          : `Không thể sao chép PNG: ${message}`,
      );
    } finally {
      pngClipboardBusyRef.current = false;
    }
  };

  const saveDocument = () => {
    historyManager.commitPendingTransaction();
    void historyManager.createCheckpoint("save-document");
    const document = toDocument();
    const blob = new Blob([JSON.stringify(document, null, 2)], {
      type: "application/json",
    });
    const url = URL.createObjectURL(blob);
    const link = window.document.createElement("a");
    const safeTitle =
      title.trim().replace(/[\\/:*?"<>|]/g, "-") ||
      (isEn ? "Untitled Formula" : "Công thức không có tiêu đề");
    link.href = url;
    link.download = safeTitle + ".visualtex.json";
    link.click();
    URL.revokeObjectURL(url);
    setSavedPulse(true);
    setToast(isEn ? "Formula document saved" : "Đã lưu tài liệu công thức");
    window.setTimeout(() => setSavedPulse(false), 900);
  };

  const openDocument = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      const parsed = JSON.parse(await file.text()) as FormulaDocument;
      if (!parsed.formulas || !Array.isArray(parsed.formulas)) {
        throw new Error("invalid");
      }
      historyManager.commitPendingTransaction();
      const before = captureDocumentSnapshot();
      loadDocument(parsed);
      const after = getEditorDocumentSnapshot({});
      if (!documentSnapshotsEquivalent(before, after)) {
        historyManager.push({
          type: "replace-document",
          before,
          after,
          source: "open-document",
          timestamp: Date.now(),
        });
        window.requestAnimationFrame(() => restoreSnapshotFocus(after));
      }
      setToast(isEn ? "Formula document opened" : "Đã mở tài liệu công thức");
    } catch {
      setToast(
        isEn
          ? "Unable to open: invalid file format"
          : "Không mở được: định dạng file không hợp lệ",
      );
    } finally {
      event.target.value = "";
    }
  };

  const newFormula = () => {
    addHistory(latex);
    const after = createBlankDocumentSnapshot(
      isEn ? "Untitled Formula" : "Công thức không có tiêu đề",
    );
    replaceDocumentWithHistory(after, "new-document");
    setToast(isEn ? "Created a blank formula" : "Tạo công thức trống");
  };

  const handleTitleChange = (nextTitle: string) => {
    const beforeTitle = useEditorStore.getState().title;
    setTitle(nextTitle);
    historyManager.recordTitleEdit({
      beforeTitle,
      afterTitle: nextTitle,
    });
  };

  const runMenuAction = (action: () => void) => {
    setMenuOpen(false);
    action();
  };

  const runUpdateCheck = useCallback(async (manual = true) => {
    if (manual) {
      setAutomaticUpdatePrompt(false);
      setUpdateResult(null);
      setUpdateOpen(true);
    }
    setUpdateChecking(true);
    setUpdateError("");
    try {
      const result = await checkForUpdates();
      setUpdateResult(result);
      if (manual || result.updateAvailable) {
        setAutomaticUpdatePrompt(!manual && result.updateAvailable);
        setUpdateOpen(true);
      }
    } catch (error) {
      if (manual) {
        setUpdateError(
          error instanceof Error
            ? error.message
            : isEn
              ? "Unable to connect to the update server"
              : "Không thể kết nối với máy chủ cập nhật",
        );
        setUpdateOpen(true);
      } else {
        automaticUpdateCheckRef.current = false;
      }
    } finally {
      setUpdateChecking(false);
    }
  }, [isEn]);

  useEffect(() => {
    if (
      !checkUpdatesOnStartup ||
      automaticUpdateCheckRef.current
    ) {
      return;
    }

    let timer = 0;
    const runWhenOnline = () => {
      if (
        automaticUpdateCheckRef.current ||
        !useEditorStore.getState().checkUpdatesOnStartup
      ) {
        return;
      }
      automaticUpdateCheckRef.current = true;
      timer = window.setTimeout(() => void runUpdateCheck(false), 1200);
    };

    window.addEventListener("online", runWhenOnline);
    if (navigator.onLine) runWhenOnline();

    return () => {
      window.clearTimeout(timer);
      window.removeEventListener("online", runWhenOnline);
    };
  }, [checkUpdatesOnStartup, runUpdateCheck]);

  useEffect(() => {
    const handleWindowKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        setMenuOpen(false);
        setCopyMenuOpen(false);
        return;
      }

      if (
        settingsOpen ||
        formulaHotkeyManagerOpen ||
        ocrOpen ||
        historyOpen ||
        exportOpen ||
        updateOpen
      ) {
        return;
      }

      const target = event.target instanceof Element ? event.target : null;
      const inCodeMirror = Boolean(target?.closest(".cm-editor"));
      const primaryModifier = (event.metaKey || event.ctrlKey) && !event.altKey;
      const key = event.key.toLowerCase();
      const requestsUndo = primaryModifier && key === "z" && !event.shiftKey;
      const requestsRedo =
        primaryModifier &&
        ((key === "z" && event.shiftKey) ||
          (key === "y" && !event.shiftKey));

      if (requestsUndo || requestsRedo) {
        if (inCodeMirror) return;
        event.preventDefault();
        if (requestsRedo) void historyManager.redo();
        else void historyManager.undo();
        return;
      }

      if (!primaryModifier) return;
      if (key === "n") {
        event.preventDefault();
        newFormula();
      } else if (key === "o") {
        event.preventDefault();
        fileInputRef.current?.click();
      } else if (key === "s") {
        event.preventDefault();
        if (keypadMode) void handleKeypadCopy();
        else saveDocument();
      } else if (key === ",") {
        event.preventDefault();
        setSettingsOpen(true);
      } else if (key === "0") {
        event.preventDefault();
        setZoom(1);
      } else if (key === "=" || key === "+") {
        event.preventDefault();
        setZoom(zoom + EDITOR_ZOOM_STEP);
      } else if (key === "-") {
        event.preventDefault();
        setZoom(zoom - EDITOR_ZOOM_STEP);
      }
    };

    window.addEventListener("keydown", handleWindowKeyDown);
    return () => window.removeEventListener("keydown", handleWindowKeyDown);
  }, [
    latex,
    title,
    isEn,
    zoom,
    keypadMode,
    keypadMinimizeOnCopy,
    latexCodeFormat,
    settingsOpen,
    formulaHotkeyManagerOpen,
    ocrOpen,
    historyOpen,
    exportOpen,
    updateOpen,
  ]);

  const codeFormatControl = (
    <div
      className={
        "copy-control code-format-control" +
        (keypadMode ? " editor-code-format-control" : "")
      }
    >
      <button
        type="button"
        className="copy-primary code-format-primary"
        aria-expanded={copyMenuOpen}
        aria-haspopup="menu"
        aria-controls="copy-format-menu"
        title={
          isEn
            ? `Current: ${currentCodeFormat.titleEn}`
            : `Hiện tại: ${currentCodeFormat.titleEn}`
        }
        onClick={() => {
          setMenuOpen(false);
          setCopyMenuOpen((open) => !open);
        }}
      >
        <Code2 size={15} />
        <span>
          {keypadMode
            ? isEn
              ? currentCodeFormat.titleEn
              : currentCodeFormat.titleVi
            : isEn
              ? "LaTeX code format"
              : "Định dạng mã LaTeX"}
        </span>
      </button>
      <button
        ref={copyMenuButtonRef}
        type="button"
        className="copy-chevron"
        aria-label={isEn ? "Choose LaTeX code format" : "Chọn định dạng mã LaTeX"}
        aria-expanded={copyMenuOpen}
        aria-haspopup="menu"
        aria-controls="copy-format-menu"
        onClick={() => {
          setMenuOpen(false);
          setCopyMenuOpen((open) => !open);
        }}
      >
        <ChevronDown size={14} />
      </button>
      {copyMenuOpen && (
        <div
          ref={copyMenuRef}
          id="copy-format-menu"
          className="copy-menu code-format-menu"
          role="menu"
          aria-label={isEn ? "LaTeX code format" : "Định dạng mã LaTeX"}
        >
          <div className="code-format-menu-header">
            <span className="copy-menu-label">
              {isEn ? "LaTeX code format" : "Định dạng mã LaTeX"}
            </span>
            <small>
              {isEn
                ? "Controls source rendering and keypad copy output"
                : "Kiểm soát kết xuất nguồn và đầu ra sao chép bàn phím"}
            </small>
          </div>
          {codeFormatGroups.map((group) => (
            <div
              className="code-format-group"
              role="group"
              aria-label={group.title}
              key={group.id}
            >
              <div className="code-format-group-heading">
                <strong>{group.title}</strong>
                <small>{group.description}</small>
              </div>
              {group.formats.map((format) => {
                const selected = format.id === latexCodeFormat;
                return (
                  <button
                    type="button"
                    role="menuitemradio"
                    aria-checked={selected}
                    aria-label={`${isEn ? format.titleEn : format.titleVi}: ${format.hint}`}
                    data-format={format.id}
                    className={selected ? "is-selected" : ""}
                    key={format.id}
                    onClick={() => handleCodeFormatChange(format.id)}
                  >
                    <span className="code-format-item-copy">
                      <small className="code-format-hint">{format.hint}</small>
                    </span>
                    <span className="code-format-check" aria-hidden="true">
                      {selected && <Check size={14} />}
                    </span>
                  </button>
                );
              })}
            </div>
          ))}
        </div>
      )}
    </div>
  );

  const keypadToggleControl = (
    <button
      type="button"
      className={`keypad-mode-toggle${keypadMode ? " is-active" : ""}`}
      aria-pressed={keypadMode}
      data-keypad-mode-toggle
      title={
        keypadMode
          ? isEn
            ? "Exit keypad mode"
            : "Thoát khỏi chế độ bàn phím"
          : isEn
            ? "Enter keypad mode"
            : "Vào chế độ bàn phím"
      }
      onClick={() => void handleKeypadModeToggle()}
    >
      <Keyboard size={15} />
      <span>{isEn ? "Keypad" : "Bàn phím"}</span>
    </button>
  );

  const desktopHeaderControls = (
    <>
      {keypadMode ? codeFormatControl : null}
      {keypadToggleControl}
    </>
  );

  return (
    <div className={`app-shell${keypadMode ? " is-keypad-mode" : ""}`}>
      <input
        ref={fileInputRef}
        type="file"
        accept=".json,.visualtex"
        className="visually-hidden"
        onChange={openDocument}
      />

      {!keypadMode && (
      <header
        className={
          "app-header" + (menuOpen || copyMenuOpen ? " has-open-menu" : "")
        }
      >
        <div className="brand-area">
          <button
            ref={menuButtonRef}
            type="button"
            className={"menu-button " + (menuOpen ? "is-active" : "")}
            aria-label={isEn ? "Main menu" : "Menu chính"}
            aria-expanded={menuOpen}
            aria-haspopup="menu"
            aria-controls="app-main-menu"
            onClick={() => {
              setCopyMenuOpen(false);
              setMenuOpen((open) => !open);
            }}
          >
            <Menu size={18} />
          </button>
          <button
            type="button"
            className={"icon-button sidebar-toggle " + (sidebarOpen ? "is-active" : "")}
            aria-label={
              editorLayout === "classic"
                ? sidebarOpen
                  ? isEn
                    ? "Hide formula tiles"
                    : "Ẩn ô công thức"
                  : isEn
                    ? "Show formula tiles"
                    : "Hiển thị ô công thức"
                : sidebarOpen
                  ? isEn
                    ? "Hide formula tools"
                    : "Ẩn công cụ công thức"
                  : isEn
                    ? "Show formula tools"
                    : "Hiện công cụ công thức"
            }
            aria-pressed={sidebarOpen}
            onClick={() => setSidebarOpen((open) => !open)}
          >
            {editorLayout === "classic" ? (
              sidebarOpen ? <PanelRightClose size={17} /> : <PanelRightOpen size={17} />
            ) : sidebarOpen ? (
              <PanelLeftClose size={17} />
            ) : (
              <PanelLeftOpen size={17} />
            )}
          </button>
          <div className="brand-mark" aria-hidden="true">
            <VisualTeXLogo className="visualtex-brand-logo" />
          </div>
          <div className="brand-copy">
            <strong>VisualTeX</strong>
          </div>

          {menuOpen && (
            <div
              ref={appMenuRef}
              id="app-main-menu"
              className="app-menu-popover"
              role="menu"
              aria-label={isEn ? "VisualTeX menu" : "Trình đơn VisualTeX"}
            >
              <div className="app-menu-heading">
                <strong>VisualTeX</strong>
                <span>{isEn ? "Formula workspace" : "Không gian làm việc công thức"}</span>
              </div>
              <button type="button" role="menuitem" onClick={() => runMenuAction(newFormula)}>
                <FilePlus2 size={16} />
                <span>{isEn ? "New formula" : "Công thức mới"}</span>
                <kbd>Ctrl+N</kbd>
              </button>
              <button
                type="button"
                role="menuitem"
                onClick={() =>
                  runMenuAction(() => fileInputRef.current?.click())
                }
              >
                <FolderOpen size={16} />
                <span>{isEn ? "Open document" : "Mở tài liệu"}</span>
                <kbd>Ctrl+O</kbd>
              </button>
              <button type="button" role="menuitem" onClick={() => runMenuAction(saveDocument)}>
                <Save size={16} />
                <span>{isEn ? "Save document" : "Lưu tài liệu"}</span>
                <kbd>Ctrl+S</kbd>
              </button>
              <button
                type="button"
                role="menuitem"
                onClick={() => runMenuAction(() => setExportOpen(true))}
              >
                <FileDown size={16} />
                <span>{isEn ? "Export…" : "Xuất…"}</span>
                <kbd>MD/SVG/PNG</kbd>
              </button>
              <div className="app-menu-divider" />
              <button
                type="button"
                role="menuitem"
                onClick={() => runMenuAction(() => setHistoryOpen(true))}
              >
                <History size={16} />
                <span>{isEn ? "Formula history" : "Lịch sử công thức"}</span>
              </button>
              <button
                type="button"
                role="menuitem"
                onClick={() => runMenuAction(() => setOcrOpen(true))}
              >
                <ScanLine size={16} />
                <span>{isEn ? "Formula image OCR" : "Hình ảnh công thức OCR"}</span>
              </button>
              <button
                type="button"
                role="menuitem"
                onClick={() => runMenuAction(() => setSettingsOpen(true))}
              >
                <Settings2 size={16} />
                <span>{isEn ? "Settings" : "Cài đặt"}</span>
              </button>
              <button
                type="button"
                role="menuitem"
                onClick={() => runMenuAction(() => setHelpOpen(true))}
              >
                <BookOpenText size={16} />
                <span>{isEn ? "Help Manual" : "Hướng dẫn trợ giúp"}</span>
              </button>
              <button
                type="button"
                role="menuitem"
                onClick={() => runMenuAction(() => void runUpdateCheck(true))}
              >
                <RefreshCw size={16} />
                <span>{isEn ? "Check for updates" : "Kiểm tra cập nhật"}</span>
              </button>
              <div className="app-menu-divider" />
              <div className="app-menu-language">
                <span>
                  <Languages size={15} />
                  {isEn ? "Language" : "Ngôn ngữ"}
                </span>
                <div>
                  <button
                    type="button"
                    role="menuitemradio"
                    aria-checked={language === "vi"}
                    className={language === "vi" ? "is-active" : ""}
                    onClick={() => setLanguage("vi")}
                  >
                    CN
                  </button>
                  <button
                    type="button"
                    role="menuitemradio"
                    aria-checked={language === "en"}
                    className={language === "en" ? "is-active" : ""}
                    onClick={() => setLanguage("en")}
                  >
                    ENG
                  </button>
                </div>
              </div>
            </div>
          )}
        </div>

        <div className="document-title-area">
          <input
            value={title}
            onChange={(event) => handleTitleChange(event.target.value)}
            onBlur={() => historyManager.commitPendingTransaction()}
            aria-label={isEn ? "Formula document title" : "Tiêu đề tài liệu công thức"}
          />
          <span
            className={"save-state " + (savedPulse ? "is-saved" : "")}
            aria-label={isEn ? "Saved" : "Đã lưu"}
            title={isEn ? "Saved" : "Đã lưu"}
          >
            <Check size={13} />
          </span>
        </div>

        <div className="header-actions">
          <div className="action-group file-actions">
            <button type="button" className="icon-button" onClick={newFormula} aria-label={isEn ? "New" : "Mới"} title={isEn ? "New · Ctrl+N" : "Mới · Ctrl+N"}>
              <FilePlus2 size={17} />
            </button>
            <button type="button" className="icon-button" onClick={() => fileInputRef.current?.click()} aria-label={isEn ? "Open" : "Mở"} title={isEn ? "Open · Ctrl+O" : "Mở · Ctrl+O"}>
              <FolderOpen size={17} />
            </button>
            <button type="button" className="icon-button" onClick={saveDocument} aria-label={isEn ? "Save" : "Lưu"} title={isEn ? "Save · Ctrl+S" : "Lưu · Ctrl+S"}>
              <Save size={17} />
            </button>
          </div>
          <div className="action-group edit-actions">
            <button
              type="button"
              className="icon-button"
              onClick={() => void historyManager.undo()}
              disabled={
                editorHistoryBusy ||
                !historyState.canUndo ||
                historyState.isReplaying
              }
              aria-label={isEn ? "Undo" : "Hoàn tác"}
              title={isEn ? "Undo · Ctrl+Z" : "Hoàn tác · Ctrl+Z"}
            >
              <Undo2 size={17} />
            </button>
            <button
              type="button"
              className="icon-button"
              onClick={() => void historyManager.redo()}
              disabled={
                editorHistoryBusy ||
                !historyState.canRedo ||
                historyState.isReplaying
              }
              aria-label={isEn ? "Redo" : "Làm lại"}
              title={isEn ? "Redo · Ctrl+Y / Ctrl+Shift+Z" : "Làm lại · Ctrl+Y / Ctrl+Shift+Z"}
            >
              <Redo2 size={17} />
            </button>
          </div>
          <button type="button" className="icon-button workspace-action" onClick={() => setHistoryOpen(true)} aria-label={isEn ? "Formula history" : "Lịch sử công thức"} title={isEn ? "Formula history" : "Lịch sử công thức"}>
            <History size={17} />
          </button>
          <button type="button" className="icon-button workspace-action" onClick={() => setOcrOpen(true)} aria-label={isEn ? "Recognize formula image" : "Nhận dạng hình ảnh công thức"} title={isEn ? "Recognize formula image" : "Nhận dạng hình ảnh công thức"}>
            <ScanLine size={17} />
          </button>
          <button type="button" className="icon-button settings-toggle" onClick={() => setSettingsOpen(true)} aria-label={isEn ? "Settings" : "Cài đặt"} title={isEn ? "Settings · Ctrl+," : "Cài đặt · Ctrl+,"}>
            <Settings2 size={17} />
          </button>
          <div className="copy-control code-format-control">
            <button
              type="button"
              className="copy-primary code-format-primary"
              aria-expanded={copyMenuOpen}
              aria-haspopup="menu"
              aria-controls="copy-format-menu"
              title={
                isEn
                  ? `Current: ${currentCodeFormat.titleEn}`
                  : `Hiện tại: ${currentCodeFormat.titleEn}`
              }
              onClick={() => {
                setMenuOpen(false);
                setCopyMenuOpen((open) => !open);
              }}
            >
              <Code2 size={16} />
              <span>{isEn ? "LaTeX code format" : "Định dạng mã LaTeX"}</span>
            </button>
            <button
              ref={copyMenuButtonRef}
              type="button"
              className="copy-chevron"
              aria-label={isEn ? "Choose LaTeX code format" : "Chọn định dạng mã LaTeX"}
              aria-expanded={copyMenuOpen}
              aria-haspopup="menu"
              aria-controls="copy-format-menu"
              onClick={() => {
                setMenuOpen(false);
                setCopyMenuOpen((open) => !open);
              }}
            >
              <ChevronDown size={15} />
            </button>
            {copyMenuOpen && (
              <div
                ref={copyMenuRef}
                id="copy-format-menu"
                className="copy-menu code-format-menu"
                role="menu"
                aria-label={isEn ? "LaTeX code format" : "Định dạng mã LaTeX"}
              >
                <div className="code-format-menu-header">
                  <span className="copy-menu-label">
                    {isEn ? "LaTeX code format" : "Định dạng mã LaTeX"}
                  </span>
                  <small>
                    {isEn
                      ? "Changes the source panel and copy output"
                      : "Thay đổi bảng nguồn và đầu ra sao chép"}
                  </small>
                </div>
                {codeFormatGroups.map((group) => (
                  <div
                    className="code-format-group"
                    role="group"
                    aria-label={group.title}
                    key={group.id}
                  >
                    <div className="code-format-group-heading">
                      <strong>{group.title}</strong>
                      <small>{group.description}</small>
                    </div>
                    {group.formats.map((format) => {
                      const selected = format.id === latexCodeFormat;
                      return (
                        <button
                          type="button"
                          role="menuitemradio"
                          aria-checked={selected}
                          aria-label={`${isEn ? format.titleEn : format.titleVi}: ${format.hint}`}
                          data-format={format.id}
                          className={selected ? "is-selected" : ""}
                          key={format.id}
                          onClick={() => handleCodeFormatChange(format.id)}
                        >
                          <span className="code-format-item-copy">
                            <small className="code-format-hint">{format.hint}</small>
                          </span>
                          <span className="code-format-check" aria-hidden="true">
                            {selected && <Check size={14} />}
                          </span>
                        </button>
                      );
                    })}
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>
      </header>
      )}

      {(menuOpen || (!keypadMode && copyMenuOpen)) && (
        <button
          type="button"
          className="menu-dismiss-layer"
          aria-label={isEn ? "Close menu" : "Đóng menu"}
          onClick={() => {
            setMenuOpen(false);
            setCopyMenuOpen(false);
          }}
        />
      )}

      <EditorWorkspace
        mode="desktop"
        showFileActions
        desktopHeaderControls={desktopHeaderControls}
        keypadMode={keypadMode}
        showUpdateActions
        showOfficeActions={false}
        showOcrActions
        onOpenExport={() => setExportOpen(true)}
        editorRef={editorRef}
        sidebarOpen={sidebarOpen}
        onSidebarOpenChange={setSidebarOpen}
        onHistoryBusyChange={setEditorHistoryBusy}
        onPasteImage={handleEditorImagePaste}
        onCopyPng={handleCopyPng}
        onCopy={async () => {
          await handleCopy();
        }}
        onReplaceDocument={replaceDocumentWithHistory}
        ocrModel={ocrModel}
        ocrModels={OCR_MODELS}
        ocrBusy={inlineOcrIsBusy || quickOcrCaptureBusy}
        onOcrModelChange={(model) =>
          handleOcrModelChange(model as OcrModelName)
        }
        onQuickOcr={() => void handleQuickOcr()}
        quickOcrCaptureMode={quickOcrCaptureMode}
        onQuickOcrCaptureModeChange={handleQuickOcrCaptureModeChange}
        silentOcrEnabled={silentOcrEnabled}
        silentOcrShortcut={silentOcrShortcut}
        onSilentOcrEnabledChange={handleSilentOcrEnabledChange}
        ocrOverlay={
          inlineOcr ? (
            <div
              className={`inline-ocr-progress is-${inlineOcr.status}`}
              role="status"
              aria-live="polite"
            >
              <span className="inline-ocr-progress-icon">
                {inlineOcr.status === "running" ||
                inlineOcr.status === "cancelling" ? (
                  <LoaderCircle size={17} className="is-spinning" />
                ) : inlineOcr.status === "success" ? (
                  <Check size={17} />
                ) : inlineOcr.status === "error" ? (
                  <AlertCircle size={17} />
                ) : (
                  <X size={17} />
                )}
              </span>
              <div>
                <strong>{inlineOcr.message}</strong>
                <span>
                  {isEn ? inlineOcrModel.labelEn : inlineOcrModel.labelVi}
                  {" · "}
                  {inlineOcr.seconds}
                  {isEn ? "s" : "s"}
                </span>
              </div>
              {inlineOcrIsBusy ? (
                <button
                  type="button"
                  className="inline-ocr-cancel"
                  onClick={cancelInlineOcr}
                  disabled={inlineOcr.status === "cancelling"}
                >
                  <X size={13} />
                  {isEn ? "Cancel" : "Hủy bỏ"}
                </button>
              ) : (
                <button
                  type="button"
                  className="inline-ocr-dismiss"
                  onClick={() => setInlineOcr(null)}
                  aria-label={isEn ? "Dismiss OCR status" : "Loại bỏ trạng thái OCR"}
                >
                  <X size={13} />
                </button>
              )}
            </div>
          ) : null
        }
      />

      <ExportDialog
        open={exportOpen}
        title={title}
        formulas={lines.map((line) => line.latex)}
        language={language}
        onClose={() => setExportOpen(false)}
        onNotify={setToast}
      />
      <HelpDialog
        open={helpOpen}
        language={language}
        onClose={() => setHelpOpen(false)}
      />
      <SettingsDialog
        open={settingsOpen}
        onClose={() => setSettingsOpen(false)}
        onCheckForUpdates={() => {
          setSettingsOpen(false);
          void runUpdateCheck(true);
        }}
        onOpenFormulaHotkeys={() => {
          setSettingsOpen(false);
          setFormulaHotkeyManagerOpen(true);
        }}
      />
      <FormulaHotkeyManagerDialog
        open={formulaHotkeyManagerOpen}
        onClose={() => setFormulaHotkeyManagerOpen(false)}
      />
      <HistoryPanel
        open={historyOpen}
        onClose={() => setHistoryOpen(false)}
        onRestore={(value) => {
          const values = value
            .replace(/\r\n?/g, "\n")
            .split("\n")
            .map(normalizeChineseLatex);
          const nextLines = reconcileFormulaLines(values, lines);
          const nextActiveLineId = nextLines.some(
            (line) => line.id === activeLineId,
          )
            ? activeLineId
            : nextLines[0]?.id ?? null;
          replaceDocumentWithHistory(
            {
              title,
              lines: nextLines,
              activeLineId: nextActiveLineId,
              formulaAlignment,
              selectionByLineId:
                editorRef.current?.getSelectionMap() ?? {},
            },
            "history-restore",
          );
          setHistoryOpen(false);
          setToast(isEn ? "Formula restored" : "Đã khôi phục công thức");
        }}
      />
      <OcrDialog
        open={ocrOpen}
        language={language}
        model={ocrModel}
        onModelChange={handleOcrModelChange}
        onClose={() => setOcrOpen(false)}
        onInsert={(value) => editorRef.current?.insertLatex(value, "ocr")}
        onAppend={(value) => editorRef.current?.appendLatex(value, "ocr")}
        onNotify={setToast}
      />
      <UpdateDialog
        open={updateOpen}
        language={language}
        checking={updateChecking}
        error={updateError}
        result={updateResult}
        checkOnStartup={checkUpdatesOnStartup}
        automaticPrompt={automaticUpdatePrompt}
        onCheckOnStartupChange={setCheckUpdatesOnStartup}
        onRetry={() => void runUpdateCheck(true)}
        onOpenRelease={() => {
          if (!updateResult) return;
          void openReleasePage(updateResult.releaseUrl).catch((error) => {
            setUpdateError(
              error instanceof Error
                ? error.message
                : isEn
                  ? "Unable to open the download page"
                  : "Không mở được trang tải xuống",
            );
          });
        }}
        onOpenProject={() => {
          void openReleasePage(PROJECT_URL).catch((error) => {
            setUpdateError(
              error instanceof Error
                ? error.message
                : isEn
                  ? "Unable to open the project page"
                  : "Không thể mở trang dự án",
            );
          });
        }}
        onClose={() => setUpdateOpen(false)}
      />

      {historyOpen && (
        <div className="panel-backdrop" onClick={() => setHistoryOpen(false)} />
      )}
      {toast && (
        <div className="toast">
          <Check size={15} />
          {toast}
        </div>
      )}
    </div>
  );
}

export default App;
