import { getCurrentWebview } from "@tauri-apps/api/webview";
import { open as openDialog } from "@tauri-apps/plugin-dialog";
import {
  AlertCircle,
  Check,
  CheckCircle2,
  ClipboardPaste,
  Copy,
  Cpu,
  Download,
  ImagePlus,
  LoaderCircle,
  Plus,
  RefreshCw,
  ScanLine,
  Trash2,
  Upload,
  X,
} from "lucide-react";
import {
  type ChangeEvent,
  type DragEvent,
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import { MathPreview } from "./MathPreview";
import {
  DEFAULT_OCR_MODEL,
  OCR_MODELS,
  cancelOcrRecognition,
  type OcrInstallProgress,
  type OcrModelName,
  type OcrRecognitionProgress,
  type OcrRecognitionResult,
  type OcrRuntimeStatus,
  fileToOcrRequest,
  getOcrRuntimeStatus,
  installOcrRuntime,
  installOptionalOcrModel,
  isOfficeCompanionEnvironment,
  isTauriEnvironment,
  listenOcrInstallProgress,
  listenOcrRecognitionProgress,
  recognizeFormulaImage,
  resolveAvailableOcrModel,
  prewarmOcrModel,
  removeOptionalOcrModel,
  resetOcrRuntime,
  restartOcrWorker,
  validateOcrImage,
} from "../ocr/ocrService";

interface OcrDialogProps {
  open: boolean;
  language: "vi" | "en";
  model: OcrModelName;
  onModelChange: (model: OcrModelName) => void;
  onClose: () => void;
  onInsert: (latex: string) => void;
  onAppend: (latex: string) => void;
  onNotify: (message: string) => void;
}

function readableBytes(bytes: number) {
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
  if (bytes < 1024 * 1024 * 1024) return (bytes / 1024 / 1024).toFixed(1) + " MB";
  return (bytes / 1024 / 1024 / 1024).toFixed(2) + " GB";
}

function readError(error: unknown) {
  if (error instanceof Error) return error.message;
  if (typeof error === "string") return error;
  try {
    return JSON.stringify(error);
  } catch {
    return "Unknown OCR error";
  }
}

const OCR_MODEL_PACKAGE_EXTENSION = ".vtxocrmodel";

function isOcrModelPackagePath(path: string) {
  return path.trim().toLowerCase().endsWith(OCR_MODEL_PACKAGE_EXTENSION);
}

function modelNameFromPackagePath(path: string): OcrModelName | null {
  const normalized = path.replaceAll("\\", "/").toLowerCase();
  return (
    OCR_MODELS.find((candidate) =>
      normalized.includes(candidate.id.toLowerCase()),
    )?.id ?? null
  );
}

function isDropPositionInsideElement(
  position: { x: number; y: number },
  element: HTMLElement | null,
) {
  if (!element) return false;
  const scale = window.devicePixelRatio || 1;
  const x = position.x / scale;
  const y = position.y / scale;
  const bounds = element.getBoundingClientRect();
  return x >= bounds.left && x <= bounds.right && y >= bounds.top && y <= bounds.bottom;
}

function normalizeResultLatex(value: string) {
  return value
    .replace(/\r\n?/g, "\n")
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean)
    .join("\n");
}

export function OcrDialog({
  open,
  language,
  model,
  onModelChange,
  onClose,
  onInsert,
  onAppend,
  onNotify,
}: OcrDialogProps) {
  const isEn = language === "en";
  const dialogRef = useRef<HTMLElement>(null);
  const previousFocusRef = useRef<HTMLElement | null>(null);
  const modelDropZoneRef = useRef<HTMLDivElement>(null);
  const modelDragPathsRef = useRef<string[]>([]);
  const recognizingRef = useRef(false);
  const cancellingRef = useRef(false);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const objectUrlRef = useRef<string | null>(null);
  const [runtime, setRuntime] = useState<OcrRuntimeStatus | null>(null);
  const [modelBusy, setModelBusy] = useState(false);
  const [modelPackageDragging, setModelPackageDragging] = useState(false);
  const [checkingRuntime, setCheckingRuntime] = useState(false);
  const [installing, setInstalling] = useState(false);
  const [installProgress, setInstallProgress] = useState<OcrInstallProgress | null>(null);
  const [file, setFile] = useState<File | null>(null);
  const [previewUrl, setPreviewUrl] = useState("");
  const [imageSize, setImageSize] = useState({ width: 0, height: 0 });
  const [dragging, setDragging] = useState(false);
  const [recognizing, setRecognizing] = useState(false);
  const [cancelling, setCancelling] = useState(false);
  const [recognitionSeconds, setRecognitionSeconds] = useState(0);
  const [recognitionProgress, setRecognitionProgress] =
    useState<OcrRecognitionProgress | null>(null);
  const [result, setResult] = useState<OcrRecognitionResult | null>(null);
  const [latex, setLatex] = useState("");
  const [error, setError] = useState("");
  const [copied, setCopied] = useState(false);

  const selectedModel = useMemo(
    () =>
      OCR_MODELS.find((item) => item.id === model) ??
      OCR_MODELS.find((item) => item.id === DEFAULT_OCR_MODEL)!,
    [model],
  );
  const defaultModel = runtime?.defaultModel ?? DEFAULT_OCR_MODEL;
  const installedModels = runtime?.installedModels ?? [];
  const selectedModelInstalled = installedModels.includes(model);
  const optionalModelMissing = model !== defaultModel && !selectedModelInstalled;
  const selectedModelRemovable =
    selectedModelInstalled && model !== defaultModel && isTauriEnvironment();

  const clearObjectUrl = useCallback(() => {
    if (objectUrlRef.current) {
      URL.revokeObjectURL(objectUrlRef.current);
      objectUrlRef.current = null;
    }
  }, []);

  const refreshRuntime = useCallback(async (forceRefresh = false) => {
    if (!isTauriEnvironment() && !isOfficeCompanionEnvironment()) {
      setRuntime({
        installed: false,
        pythonPath: null,
        pythonVersion: null,
        paddleVersion: null,
        paddleocrVersion: null,
        runtimePath: "",
        offlineBundleAvailable: false,
        installedModels: [],
        defaultModel: DEFAULT_OCR_MODEL,
        message: isEn
          ? "OCR is available in the VisualTeX desktop app, not in the browser preview."
          : "OCR có sẵn trong ứng dụng VisualTeX dành cho máy tính để bàn, không có trong bản xem trước của trình duyệt.",
      });
      return;
    }

    setCheckingRuntime(true);
    try {
      setRuntime(await getOcrRuntimeStatus(forceRefresh));
    } catch (runtimeError) {
      setError(readError(runtimeError));
    } finally {
      setCheckingRuntime(false);
    }
  }, [isEn]);

  const importModelPackage = useCallback(
    async (packagePath: string) => {
      if (!isTauriEnvironment() || modelBusy) return;
      if (!isOcrModelPackagePath(packagePath)) {
        setError(
          isEn
            ? "Choose a VisualTeX .vtxocrmodel package."
            : "Chọn gói VisualTeX .vtxocrmodel.",
        );
        return;
      }

      const previouslyInstalled = new Set(runtime?.installedModels ?? []);
      const packageModel = modelNameFromPackagePath(packagePath);
      setModelBusy(true);
      setError("");
      try {
        const nextRuntime = await installOptionalOcrModel(packagePath);
        setRuntime(nextRuntime);
        const newlyInstalled = nextRuntime.installedModels.find(
          (candidate) => !previouslyInstalled.has(candidate),
        ) as OcrModelName | undefined;
        const imported =
          newlyInstalled ??
          (packageModel && nextRuntime.installedModels.includes(packageModel)
            ? packageModel
            : undefined) ??
          (nextRuntime.installedModels.includes(model) ? model : undefined) ??
          (nextRuntime.installedModels.at(-1) as OcrModelName | undefined);
        if (imported) onModelChange(imported);
        onNotify(isEn ? "Verified OCR model imported" : "Đã nhập mô hình OCR đã được xác minh");
      } catch (importError) {
        setError(readError(importError));
      } finally {
        setModelBusy(false);
      }
    },
    [isEn, model, modelBusy, onModelChange, onNotify, runtime?.installedModels],
  );

  useEffect(() => {
    if (!open) return;
    setError("");
    if (runtime) return;

    let cancelled = false;
    const frame = window.requestAnimationFrame(() => {
      if (!cancelled) void refreshRuntime();
    });
    return () => {
      cancelled = true;
      window.cancelAnimationFrame(frame);
    };
  }, [open, runtime, refreshRuntime]);

  useEffect(() => {
    if (!open || !isTauriEnvironment()) return;
    let unlisten: (() => void) | undefined;
    let disposed = false;

    const updateDropHighlight = (
      paths: string[],
      position: { x: number; y: number },
    ) => {
      setModelPackageDragging(
        !modelBusy &&
          paths.some(isOcrModelPackagePath) &&
          isDropPositionInsideElement(position, modelDropZoneRef.current),
      );
    };

    void getCurrentWebview()
      .onDragDropEvent((event) => {
        if (disposed) return;
        const payload = event.payload;
        if (payload.type === "enter") {
          modelDragPathsRef.current = payload.paths;
          updateDropHighlight(payload.paths, payload.position);
          return;
        }
        if (payload.type === "over") {
          updateDropHighlight(modelDragPathsRef.current, payload.position);
          return;
        }
        if (payload.type === "leave") {
          modelDragPathsRef.current = [];
          setModelPackageDragging(false);
          return;
        }

        const droppedInsideModelArea = isDropPositionInsideElement(
          payload.position,
          modelDropZoneRef.current,
        );
        const packages = payload.paths.filter(isOcrModelPackagePath);
        modelDragPathsRef.current = [];
        setModelPackageDragging(false);
        if (!droppedInsideModelArea) return;
        if (packages.length !== 1) {
          setError(
            packages.length > 1
              ? isEn
                ? "Import one .vtxocrmodel package at a time."
                : "Nhập một gói .vtxocrmodel mỗi lần."
              : isEn
                ? "Drop a VisualTeX .vtxocrmodel package here."
                : "Thả gói VisualTeX .vtxocrmodel vào đây.",
          );
          return;
        }
        void importModelPackage(packages[0]);
      })
      .then((dispose) => {
        if (disposed) dispose();
        else unlisten = dispose;
      })
      .catch((dragDropError) => {
        if (!disposed) setError(readError(dragDropError));
      });

    return () => {
      disposed = true;
      modelDragPathsRef.current = [];
      setModelPackageDragging(false);
      unlisten?.();
    };
  }, [importModelPackage, isEn, modelBusy, open]);

  useEffect(() => {
    if (!open) return;

    const handlePaste = (event: ClipboardEvent) => {
      const item = Array.from(event.clipboardData?.items ?? []).find((candidate) =>
        candidate.type.startsWith("image/"),
      );
      const pastedFile = item?.getAsFile();
      if (!pastedFile) return;
      event.preventDefault();
      try {
        validateOcrImage(pastedFile);
        clearObjectUrl();
        const nextUrl = URL.createObjectURL(pastedFile);
        objectUrlRef.current = nextUrl;
        setPreviewUrl(nextUrl);
        setFile(pastedFile);
        setImageSize({ width: 0, height: 0 });
        setResult(null);
        setLatex("");
        setError("");
      } catch (pasteError) {
        setError(readError(pasteError));
      }
    };

    window.addEventListener("paste", handlePaste);
    return () => window.removeEventListener("paste", handlePaste);
  }, [open, clearObjectUrl]);

  useEffect(
    () => () => {
      clearObjectUrl();
    },
    [clearObjectUrl],
  );

  useEffect(() => {
    if (!recognizing) return;
    setRecognitionSeconds(0);
    const startedAt = Date.now();
    const timer = window.setInterval(() => {
      setRecognitionSeconds(Math.floor((Date.now() - startedAt) / 1000));
    }, 1000);
    return () => window.clearInterval(timer);
  }, [recognizing]);

  useEffect(() => {
    recognizingRef.current = recognizing;
  }, [recognizing]);

  useEffect(() => {
    cancellingRef.current = cancelling;
  }, [cancelling]);

  const selectFile = useCallback(
    (nextFile: File) => {
      try {
        validateOcrImage(nextFile);
        clearObjectUrl();
        const nextUrl = URL.createObjectURL(nextFile);
        objectUrlRef.current = nextUrl;
        setPreviewUrl(nextUrl);
        setFile(nextFile);
        setImageSize({ width: 0, height: 0 });
        setResult(null);
        setLatex("");
        setError("");
      } catch (selectionError) {
        setError(readError(selectionError));
      }
    },
    [clearObjectUrl],
  );

  const handleFileInput = (event: ChangeEvent<HTMLInputElement>) => {
    const nextFile = event.target.files?.[0];
    if (nextFile) selectFile(nextFile);
    event.target.value = "";
  };

  const handleDrop = (event: DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setDragging(false);
    const nextFile = Array.from(event.dataTransfer.files).find((candidate) =>
      candidate.type.startsWith("image/"),
    );
    if (nextFile) selectFile(nextFile);
    else setError(isEn ? "Drop an image file here." : "Thả file ảnh vào đây.");
  };

  const handleInstall = async () => {
    setInstalling(true);
    setError("");
    setInstallProgress({
      stage: "start",
      percent: 1,
      message: isEn ? "Starting OCR installation" : "Bắt đầu cài đặt OCR",
      detail: null,
    });

    let unlisten: (() => void) | undefined;
    try {
      unlisten = await listenOcrInstallProgress(setInstallProgress);
      const nextRuntime = await installOcrRuntime();
      setRuntime(nextRuntime);
      const availableModel = resolveAvailableOcrModel(nextRuntime, model);
      void prewarmOcrModel(availableModel).catch(() => undefined);
      onNotify(isEn ? "OCR runtime installed" : "Đã cài đặt thời gian chạy OCR");
    } catch (installError) {
      setError(readError(installError));
      await refreshRuntime(true);
    } finally {
      unlisten?.();
      setInstalling(false);
    }
  };

  const handleImportModel = async () => {
    if (!isTauriEnvironment() || modelBusy) return;
    try {
      const selected = await openDialog({
        multiple: false,
        directory: false,
        filters: [
          {
            name: "VisualTeX OCR model",
            extensions: ["vtxocrmodel"],
          },
        ],
      });
      if (typeof selected === "string") await importModelPackage(selected);
    } catch (importError) {
      setError(readError(importError));
    }
  };

  const handleRemoveModel = async () => {
    if (!selectedModelRemovable || modelBusy) return;
    const confirmed = window.confirm(
      isEn
        ? `Remove the installed ${selectedModel.labelEn} model? The OCR runtime and other models will be kept.`
        : `Xóa mẫu ${selectedModel.labelEn} đã cài đặt? Thời gian chạy OCR và các mô hình khác sẽ được giữ nguyên.`,
    );
    if (!confirmed) return;
    setModelBusy(true);
    setError("");
    try {
      const nextRuntime = await removeOptionalOcrModel(model);
      setRuntime(nextRuntime);
      const fallback = resolveAvailableOcrModel(nextRuntime, DEFAULT_OCR_MODEL);
      onModelChange(fallback);
      onNotify(isEn ? "Optional OCR model removed" : "Đã xóa mô hình OCR tùy chọn");
    } catch (removeError) {
      setError(readError(removeError));
    } finally {
      setModelBusy(false);
    }
  };

  const handleRecognize = async () => {
    if (!file) {
      setError(isEn ? "Choose or paste a formula image first." : "Trước tiên hãy chọn hoặc dán hình ảnh công thức.");
      return;
    }
    if (!runtime?.installed) {
      setError(isEn ? "Install the OCR runtime first." : "Trước tiên hãy cài đặt thời gian chạy OCR.");
      return;
    }
    if (!selectedModelInstalled) {
      setError(
        isEn
          ? "Import the selected OCR model package first."
          : "Trước tiên hãy nhập gói mô hình OCR đã chọn.",
      );
      return;
    }

    setRecognizing(true);
    cancellingRef.current = false;
    setCancelling(false);
    setRecognitionProgress({
      event: "progress",
      id: "pending",
      stage: "preprocess",
      model,
      message: isEn ? "Preparing the formula image" : "Chuẩn bị hình ảnh công thức",
    });
    setResult(null);
    setLatex("");
    setError("");

    let unlisten: (() => void) | undefined;
    try {
      unlisten = await listenOcrRecognitionProgress((progress) => {
        if (progress.model === model) setRecognitionProgress(progress);
      });
      const request = await fileToOcrRequest(file, model);
      const nextResult = await recognizeFormulaImage(request);
      setResult(nextResult);
      setLatex(
        normalizeResultLatex(nextResult.formulas.map((formula) => formula.latex).join("\n")),
      );
    } catch (recognitionError) {
      const message = readError(recognitionError);
      if (cancellingRef.current || message.includes("OCR_CANCELLED")) {
        onNotify(isEn ? "OCR recognition cancelled" : "Nhận dạng OCR đã bị hủy");
      } else {
        setError(message);
      }
    } finally {
      unlisten?.();
      setRecognizing(false);
      cancellingRef.current = false;
      setCancelling(false);
      setRecognitionProgress(null);
    }
  };

  const handleCancelRecognition = async () => {
    if (!recognizing || cancelling) return;
    cancellingRef.current = true;
    setCancelling(true);
    setRecognitionProgress((current) => ({
      event: "progress",
      id: current?.id ?? "pending",
      stage: "cancelling",
      model,
      message: isEn ? "Stopping the OCR worker…" : "Đang dừng nhân viên OCR…",
    }));
    try {
      await cancelOcrRecognition();
    } catch (cancelError) {
      setError(readError(cancelError));
      cancellingRef.current = false;
      setCancelling(false);
    }
  };

  const requestClose = () => {
    if (recognizingRef.current) void handleCancelRecognition();
    onClose();
  };

  useEffect(() => {
    if (!open) return;
    previousFocusRef.current = document.activeElement as HTMLElement | null;
    const frame = window.requestAnimationFrame(() => {
      dialogRef.current?.querySelector<HTMLElement>("button, input, select")?.focus();
    });

    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.preventDefault();
        requestClose();
        return;
      }
      if (event.key !== "Tab" || !dialogRef.current) return;

      const focusable = Array.from(
        dialogRef.current.querySelectorAll<HTMLElement>(
          'button:not(:disabled), input:not(:disabled), select:not(:disabled), textarea:not(:disabled), [tabindex]:not([tabindex="-1"])',
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

  const handleCopy = async () => {
    const value = normalizeResultLatex(latex);
    if (!value) return;
    await navigator.clipboard.writeText(value);
    setCopied(true);
    window.setTimeout(() => setCopied(false), 1200);
  };

  const handleInsert = () => {
    const value = normalizeResultLatex(latex);
    if (!value) return;
    onInsert(value);
    onNotify(isEn ? "OCR formula inserted at the cursor" : "Công thức OCR được chèn vào con trỏ");
    onClose();
  };

  const handleAppend = () => {
    const value = normalizeResultLatex(latex);
    if (!value) return;
    onAppend(value);
    onNotify(isEn ? "OCR formula appended as a new line" : "Công thức OCR được thêm vào dưới dạng dòng mới");
    onClose();
  };

  const handleRestartWorker = async () => {
    try {
      await restartOcrWorker();
      await prewarmOcrModel(model);
      setResult(null);
      setLatex("");
      setError("");
      onNotify(isEn ? "OCR worker restarted" : "Nhân viên OCR đã khởi động lại");
    } catch (restartError) {
      setError(readError(restartError));
    }
  };

  const handleResetRuntime = async () => {
    const confirmed = window.confirm(
      isEn
        ? "Remove the OCR runtime and its installed packages?"
        : "Xóa thời gian chạy OCR và các gói đã cài đặt của nó?",
    );
    if (!confirmed) return;

    setCheckingRuntime(true);
    setError("");
    try {
      setRuntime(await resetOcrRuntime());
      setResult(null);
      setLatex("");
    } catch (resetError) {
      setError(readError(resetError));
    } finally {
      setCheckingRuntime(false);
    }
  };

  if (!open) return null;

  return (
    <div
      className="modal-backdrop ocr-modal-backdrop"
      role="presentation"
      onMouseDown={requestClose}
    >
      <section
        ref={dialogRef}
        className="ocr-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="ocr-dialog-title"
        onMouseDown={(event) => event.stopPropagation()}
      >
        <header className="dialog-header ocr-dialog-header">
          <div className="ocr-heading">
            <span className="ocr-heading-icon">
              <ScanLine size={20} />
            </span>
            <div>
              <span className="eyebrow">PP-FORMULANET OCR</span>
              <h2 id="ocr-dialog-title">{isEn ? "Formula image recognition" : "Nhận dạng hình ảnh công thức"}</h2>
            </div>
          </div>
          <button
            type="button"
            className="icon-button"
            onClick={requestClose}
            aria-label={isEn ? "Close OCR" : "Đóng OCR"}
          >
            <X size={18} />
          </button>
        </header>

        <div className="ocr-dialog-body">
          <div className="ocr-input-column">
            <input
              ref={fileInputRef}
              type="file"
              className="visually-hidden"
              accept="image/png,image/jpeg,image/webp,image/bmp,image/tiff"
              onChange={handleFileInput}
            />

            <div
              className={
                "ocr-drop-zone" +
                (dragging ? " is-dragging" : "") +
                (previewUrl ? " has-image" : "")
              }
              onDragEnter={(event) => {
                event.preventDefault();
                setDragging(true);
              }}
              onDragOver={(event) => event.preventDefault()}
              onDragLeave={(event) => {
                if (event.currentTarget === event.target) setDragging(false);
              }}
              onDrop={handleDrop}
            >
              {previewUrl ? (
                <>
                  <img
                    src={previewUrl}
                    alt={isEn ? "Formula source preview" : "Xem trước nguồn công thức"}
                    onLoad={(event) =>
                      setImageSize({
                        width: event.currentTarget.naturalWidth,
                        height: event.currentTarget.naturalHeight,
                      })
                    }
                  />
                  <div className="ocr-image-actions">
                    <button type="button" onClick={() => fileInputRef.current?.click()}>
                      <RefreshCw size={14} />
                      {isEn ? "Replace" : "Thay thế"}
                    </button>
                  </div>
                </>
              ) : (
                <div className="ocr-drop-empty">
                  <span className="ocr-drop-icon">
                    <ImagePlus size={28} />
                  </span>
                  <strong>{isEn ? "Drop a formula image here" : "Thả ảnh công thức vào đây"}</strong>
                  <span>{isEn ? "Choose a file or paste an image" : "Chọn file hoặc dán hình ảnh"}</span>
                  <button type="button" onClick={() => fileInputRef.current?.click()}>
                    <Upload size={15} />
                    {isEn ? "Choose image" : "Chọn hình ảnh"}
                  </button>
                  <small>
                    <ClipboardPaste size={13} />
                    {isEn ? "Paste with ⌘V while this dialog is open" : "Dán bằng ⌘V khi hộp thoại này đang mở"}
                  </small>
                </div>
              )}
            </div>

            {file && (
              <div className="ocr-file-meta">
                <span>{file.name || (isEn ? "Clipboard image" : "Hình ảnh bảng nhớ tạm")}</span>
                <span>
                  {imageSize.width > 0 ? imageSize.width + "×" + imageSize.height + " · " : ""}
                  {readableBytes(file.size)}
                </span>
              </div>
            )}

            <label className="ocr-model-field">
              <span>{isEn ? "Recognition model" : "Model nhận dạng"}</span>
              <select
                value={model}
                disabled={recognizing || cancelling || modelBusy}
                onChange={(event) =>
                  onModelChange(event.target.value as OcrModelName)
                }
              >
                {OCR_MODELS.map((item) => {
                  const available = installedModels.includes(item.id);
                  return (
                    <option value={item.id} key={item.id}>
                      {isEn ? item.labelEn : item.labelVi}
                      {available
                        ? isEn
                          ? " · installed"
                          : "· đã cài đặt"
                        : isEn
                          ? " · not installed"
                          : "· chưa được cài đặt"}
                    </option>
                  );
                })}
              </select>
              <small>{isEn ? selectedModel.hintEn : selectedModel.hintVi}</small>
            </label>

            <div
              ref={modelDropZoneRef}
              className={
                "ocr-model-warning ocr-model-drop-zone" +
                (modelPackageDragging ? " is-dragging" : "")
              }
              role="status"
            >
              {selectedModelInstalled ? (
                <CheckCircle2 size={15} />
              ) : (
                <AlertCircle size={15} />
              )}
              <div>
                <strong>
                  {selectedModelInstalled
                    ? isEn
                      ? `${selectedModel.labelEn} is installed`
                      : `${selectedModel.labelEn} đã được cài đặt`
                    : isEn
                      ? `${selectedModel.labelEn} is not installed`
                      : `${selectedModel.labelEn} chưa được cài đặt`}
                </strong>
                <span>
                  {selectedModelInstalled
                    ? isEn
                      ? "Recognition and warmup use this verified local model."
                      : "Nhận dạng và khởi động sử dụng mô hình cục bộ đã được xác minh này."
                    : isEn
                      ? "Import the matching verified .vtxocrmodel package before recognition."
                      : "Nhập gói .vtxocrmodel đã được xác minh phù hợp trước khi nhận dạng."}
                </span>
                <small className="ocr-model-drop-hint">
                  {modelPackageDragging
                    ? isEn
                      ? "Release to verify and import this model package"
                      : "Phát hành để xác minh và nhập gói mô hình này"
                    : isEn
                      ? "Choose a package below, or drag one .vtxocrmodel package into this area."
                      : "Chọn gói bên dưới hoặc kéo một gói .vtxocrmodel vào khu vực này."}
                </small>
                <div className="ocr-install-actions">
                  {!selectedModelInstalled && (
                    <button
                      type="button"
                      onClick={() => void handleImportModel()}
                      disabled={!isTauriEnvironment() || modelBusy}
                    >
                      {modelBusy ? (
                        <LoaderCircle size={14} className="is-spinning" />
                      ) : (
                        <Upload size={14} />
                      )}
                      {isEn ? "Import package" : "Gói hàng nhập khẩu"}
                    </button>
                  )}
                  {selectedModelRemovable && (
                    <button
                      type="button"
                      className="is-danger"
                      onClick={() => void handleRemoveModel()}
                      disabled={modelBusy || recognizing}
                    >
                      <Trash2 size={14} />
                      {isEn ? "Remove model" : "Xóa mô hình"}
                    </button>
                  )}
                </div>
              </div>
            </div>

            {model === "PP-FormulaNet_plus-L" && (
              <div className="ocr-model-warning" role="note">
                <AlertCircle size={15} />
                <span>
                  {isEn
                    ? "The L model occupies about 698 MB and can use several GB of memory. Use M unless L accuracy is necessary."
                    : "Model L chiếm khoảng 698 MB và có thể sử dụng vài GB bộ nhớ. Sử dụng M trừ khi cần độ chính xác L."}
                </span>
              </div>
            )}

            <div className="ocr-input-tip">
              <AlertCircle size={14} />
              <span>
                {isEn
                  ? "Use a tight crop around one formula. Avoid blur, shadows, and perspective distortion."
                  : "Sử dụng cắt xén chặt chẽ xung quanh một công thức. Tránh làm mờ, đổ bóng và biến dạng phối cảnh."}
              </span>
            </div>
          </div>

          <div className="ocr-output-column">
            <section className="ocr-runtime-card">
              <div className="ocr-runtime-summary">
                <span className={"ocr-runtime-icon " + (runtime?.installed ? "is-ready" : "")}>
                  {checkingRuntime ? (
                    <LoaderCircle size={17} className="is-spinning" />
                  ) : runtime?.installed ? (
                    <CheckCircle2 size={17} />
                  ) : (
                    <Cpu size={17} />
                  )}
                </span>
                <div>
                  <strong>
                    {runtime?.installed
                      ? isEn
                        ? "Local OCR runtime ready"
                        : "Thời gian chạy OCR cục bộ đã sẵn sàng"
                      : isEn
                        ? "OCR runtime is not installed"
                        : "Thời gian chạy OCR chưa được cài đặt"}
                  </strong>
                  <span>{runtime?.message ?? (isEn ? "Checking runtime…" : "Đang kiểm tra thời gian chạy…")}</span>
                </div>
              </div>

              {runtime?.installed ? (
                <div className="ocr-runtime-details">
                  <span>Python {runtime.pythonVersion}</span>
                  <span>Paddle {runtime.paddleVersion}</span>
                  <span>PaddleOCR {runtime.paddleocrVersion}</span>
                  <span>
                    {isEn ? "Models" : "Model"}: {installedModels.length > 0
                      ? installedModels.join(", ")
                      : isEn
                        ? "none"
                        : "không có"}
                  </span>
                  <button
                    type="button"
                    onClick={handleRestartWorker}
                    disabled={!selectedModelInstalled}
                  >
                    <RefreshCw size={13} />
                    {isEn ? "Restart" : "Khởi động lại"}
                  </button>
                  <button type="button" className="is-danger" onClick={handleResetRuntime}>
                    <Trash2 size={13} />
                    {isEn ? "Reset" : "Đặt lại"}
                  </button>
                </div>
              ) : (
                <div className="ocr-install-panel">
                  {installing && installProgress ? (
                    <>
                      <div className="ocr-progress-label">
                        <span>{installProgress.message}</span>
                        <strong>{installProgress.percent}%</strong>
                      </div>
                      <div className="ocr-progress-track">
                        <span style={{ width: installProgress.percent + "%" }} />
                      </div>
                      {installProgress.detail && <small>{installProgress.detail}</small>}
                    </>
                  ) : (
                    <>
                      <p>
                        {isEn
                          ? "VisualTeX will verify and extract the bundled Python 3.10, PaddlePaddle 3.3.1, PaddleOCR 3.7.0, and the default M model entirely on this Mac. No network or pip installation is used."
                          : "VisualTeX sẽ xác minh và trích xuất gói Python 3.10, PaddlePaddle 3.3.1, PaddleOCR 3.7.0 và mô hình M mặc định hoàn toàn trên máy Mac này. Không có cài đặt mạng hoặc pip được sử dụng."}
                      </p>
                      <button
                        type="button"
                        className="primary-button"
                        onClick={handleInstall}
                        disabled={
                          (!isTauriEnvironment() &&
                            !isOfficeCompanionEnvironment()) ||
                          checkingRuntime
                        }
                      >
                        <Download size={15} />
                        {isEn ? "Install OCR runtime" : "Cài đặt thời gian chạy OCR"}
                      </button>
                    </>
                  )}
                </div>
              )}
            </section>

            <section className="ocr-result-card">
              <div className="ocr-result-heading">
                <div>
                  <span className="eyebrow">LATEX RESULT</span>
                  <strong>{isEn ? "Recognition result" : "Kết quả nhận dạng"}</strong>
                </div>
                {result && (
                  <span>
                    {result.backgroundInverted
                      ? isEn
                        ? "Dark background normalized · "
                        : "Đã chuẩn hóa nền tối ·"
                      : ""}
                    {result.elapsedMs} ms · {result.processedWidth}×{result.processedHeight}
                  </span>
                )}
              </div>

              {recognizing ? (
                <div className="ocr-recognizing-state">
                  <LoaderCircle size={24} className="is-spinning" />
                  <strong>
                    {recognitionProgress?.message ??
                      (isEn ? "Recognizing formula…" : "Nhận biết công thức…")}
                  </strong>
                  <span>
                    {isEn
                      ? `${selectedModel.labelEn} · ${recognitionSeconds}s elapsed`
                      : `${selectedModel.labelEn} · ${recognitionSeconds} đã trôi qua`}
                  </span>
                  <small className="ocr-recognition-meta">
                    {isEn
                      ? `First use may download ${selectedModel.downloadMb.toFixed(1)} MB. You can cancel without closing VisualTeX.`
                      : `Lần sử dụng đầu tiên có thể tải xuống ${selectedModel.downloadMb.toFixed(1)} MB. Bạn có thể hủy mà không cần đóng VisualTeX.`}
                  </small>
                </div>
              ) : latex ? (
                <>
                  <div className="ocr-formula-preview">
                    <MathPreview latex={latex.split("\n")[0]} />
                  </div>
                  <label className="ocr-latex-editor">
                    <span>{isEn ? "Editable LaTeX" : "LaTeX có thể chỉnh sửa"}</span>
                    <textarea value={latex} onChange={(event) => setLatex(event.target.value)} spellCheck={false} />
                  </label>
                </>
              ) : (
                <div className="ocr-empty-result">
                  <ScanLine size={24} />
                  <span>
                    {isEn
                      ? "Choose an image and run recognition."
                      : "Chọn một hình ảnh và chạy nhận dạng."}
                  </span>
                </div>
              )}
            </section>

            {error && (
              <div className="ocr-error-box" role="alert">
                <AlertCircle size={16} />
                <pre>{error}</pre>
              </div>
            )}
          </div>
        </div>

        <footer className="dialog-footer ocr-dialog-footer">
          {recognizing ? (
            <button
              type="button"
              className="secondary-button is-danger"
              onClick={handleCancelRecognition}
              disabled={cancelling}
            >
              {cancelling ? (
                <LoaderCircle size={15} className="is-spinning" />
              ) : (
                <X size={15} />
              )}
              {cancelling
                ? isEn
                  ? "Stopping…"
                  : "Đang dừng…"
                : isEn
                  ? "Cancel recognition"
                  : "Hủy nhận dạng"}
            </button>
          ) : (
            <button
              type="button"
              className="secondary-button"
              onClick={handleRecognize}
              disabled={
                !file ||
                !runtime?.installed ||
                !selectedModelInstalled ||
                installing ||
                modelBusy
              }
            >
              <ScanLine size={15} />
              {isEn ? "Recognize" : "Nhận biết"}
            </button>
          )}
          <div className="ocr-result-actions">
            <button type="button" className="secondary-button" onClick={handleCopy} disabled={!latex.trim()}>
              {copied ? <Check size={15} /> : <Copy size={15} />}
              {copied ? (isEn ? "Copied" : "Sao chép") : isEn ? "Copy LaTeX" : "Sao chép LaTeX"}
            </button>
            <button type="button" className="secondary-button" onClick={handleAppend} disabled={!latex.trim()}>
              <Plus size={15} />
              {isEn ? "Append line" : "Nối dòng"}
            </button>
            <button type="button" className="primary-button" onClick={handleInsert} disabled={!latex.trim()}>
              <ScanLine size={15} />
              {isEn ? "Insert at cursor" : "Chèn vào con trỏ"}
            </button>
          </div>
        </footer>
      </section>
    </div>
  );
}
