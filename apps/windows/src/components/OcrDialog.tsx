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
  FolderOpen,
  HardDrive,
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
  beginOcrInstallGuard,
  endOcrInstallGuard,
  isOcrInstallActive,
  ocrInstallStatusToProgress,
  shouldDisplayRuntimeError,
} from "../ocr/ocrInstallState";
import {
  DEFAULT_OCR_MODEL,
  OCR_MODELS,
  cancelOcrInstall,
  cancelOcrModelDownload,
  cancelOcrRecognition,
  configureOcrStorageLocation,
  downloadOcrModel,
  getOcrInstallStatus,
  getOcrModelCatalog,
  getOcrModelDownloadStatus,
  openOcrInstallLogs,
  openOcrStorageLocation,
  type OcrInstallProgress,
  type OcrInstallStatus,
  type OcrModelCatalog,
  type OcrModelDownloadSnapshot,
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
  listenOcrModelDownloadProgress,
  listenOcrRecognitionProgress,
  recognizeFormulaImage,
  removeOptionalOcrModel,
  resolveAvailableOcrModel,
  resetOcrRuntime,
  restartOcrWorker,
  warmupOcrModel,
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

function readableEta(seconds: number | null) {
  if (seconds === null || !Number.isFinite(seconds)) return "";
  if (seconds < 60) return `${Math.max(0, Math.round(seconds))}s`;
  const minutes = Math.floor(seconds / 60);
  const remainder = Math.round(seconds % 60);
  return `${minutes}m ${remainder}s`;
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
    OCR_MODELS.find((candidate) => normalized.includes(candidate.id.toLowerCase()))?.id ?? null
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

function hasTauriWebviewRuntime() {
  const internals = (
    window as Window & {
      __TAURI_INTERNALS__?: { metadata?: unknown };
    }
  ).__TAURI_INTERNALS__;
  return Boolean(internals?.metadata);
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
  const installingRef = useRef(false);
  const modelCancelRequestedRef = useRef(false);
  const runtimeRequestGenerationRef = useRef(0);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const objectUrlRef = useRef<string | null>(null);
  const [runtime, setRuntime] = useState<OcrRuntimeStatus | null>(null);
  const [modelCatalog, setModelCatalog] = useState<OcrModelCatalog | null>(null);
  const [modelDownload, setModelDownload] = useState<OcrModelDownloadSnapshot | null>(null);
  const [modelBusy, setModelBusy] = useState(false);
  const [modelPackageDragging, setModelPackageDragging] = useState(false);
  const [checkingRuntime, setCheckingRuntime] = useState(false);
  const [changingStorage, setChangingStorage] = useState(false);
  const [installing, setInstalling] = useState(false);
  const [installProgress, setInstallProgress] = useState<OcrInstallProgress | null>(null);
  const [installStatus, setInstallStatus] = useState<OcrInstallStatus | null>(null);
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

  const isWindows = /Windows/i.test(navigator.userAgent);
  const installFailed =
    installStatus?.state === "installFailed" ||
    installStatus?.state === "verificationFailed" ||
    installStatus?.state === "cancelled";

  const selectedModel = useMemo(
    () =>
      OCR_MODELS.find((item) => item.id === model) ??
      OCR_MODELS.find((item) => item.id === DEFAULT_OCR_MODEL)!,
    [model],
  );
  const installedModels = runtime?.installedModels ?? [];
  const damagedModels = runtime?.damagedModels ?? [];
  const selectedModelInstalled = installedModels.includes(model);
  const selectedCatalogEntry = modelCatalog?.entries.find((entry) => entry.model === model);
  const modelDownloadActive =
    modelDownload?.state === "downloading" ||
    modelDownload?.state === "verifying" ||
    modelDownload?.state === "installing";
  const storageAvailableBytes = runtime?.storageAvailableBytes ?? null;
  const storageLowForInitialInstall =
    !runtime?.installed &&
    storageAvailableBytes !== null &&
    storageAvailableBytes < 2 * 1024 * 1024 * 1024;

  const applyRuntimeStatus = useCallback((nextRuntime: OcrRuntimeStatus) => {
    // Invalidate any older status request that may still be resolving. Without
    // this guard, an OCR dialog refresh started before a reset/path switch can
    // arrive later and overwrite the new path or uninstalled state.
    runtimeRequestGenerationRef.current += 1;
    setCheckingRuntime(false);
    setRuntime(nextRuntime);
  }, []);

  const clearObjectUrl = useCallback(() => {
    if (objectUrlRef.current) {
      URL.revokeObjectURL(objectUrlRef.current);
      objectUrlRef.current = null;
    }
  }, []);

  const applyInstallStatus = useCallback((status: OcrInstallStatus) => {
    const active = isOcrInstallActive(status.state);
    setInstallStatus(status);
    setInstallProgress(ocrInstallStatusToProgress(status));
    installingRef.current = active;
    setInstalling(active);
  }, []);

  const refreshInstallStatus = useCallback(async () => {
    if (!isTauriEnvironment() && !isOfficeCompanionEnvironment()) return;
    try {
      applyInstallStatus(await getOcrInstallStatus());
    } catch (statusError) {
      setError(readError(statusError));
    }
  }, [applyInstallStatus]);

  const refreshRuntime = useCallback(async (forceRefresh = false) => {
    const requestGeneration = ++runtimeRequestGenerationRef.current;
    if (!isTauriEnvironment() && !isOfficeCompanionEnvironment()) {
      if (requestGeneration === runtimeRequestGenerationRef.current) {
        setRuntime({
          installed: false,
          pythonPath: null,
          pythonVersion: null,
          paddleVersion: null,
          paddleocrVersion: null,
          runtimePath: "",
          storageConfigPath: "",
          storageSource: "default",
          storageManaged: false,
          storageAvailableBytes: null,
          storagePersistentAcrossUninstall: false,
          runtimeBundleAvailable: false,
          offlineBundleAvailable: false,
          installedModels: [],
          damagedModels: [],
          modelCatalogAvailable: false,
          defaultModel: "PP-FormulaNet_plus-M",
          message: isEn
            ? "OCR is available in the VisualTeX desktop app, not in the browser preview."
            : "OCR có sẵn trong ứng dụng VisualTeX dành cho máy tính để bàn, không có trong bản xem trước của trình duyệt.",
        });
      }
      return;
    }

    setCheckingRuntime(true);
    try {
      const nextRuntime = await getOcrRuntimeStatus(forceRefresh);
      if (requestGeneration === runtimeRequestGenerationRef.current) {
        setRuntime(nextRuntime);
      }
    } catch (runtimeError) {
      if (requestGeneration === runtimeRequestGenerationRef.current) {
        setError(readError(runtimeError));
      }
    } finally {
      if (requestGeneration === runtimeRequestGenerationRef.current) {
        setCheckingRuntime(false);
      }
    }
  }, [isEn]);

  const refreshModelCatalog = useCallback(async () => {
    if (!isTauriEnvironment()) {
      setModelCatalog(null);
      return;
    }
    try {
      const [catalog, downloadStatus] = await Promise.all([
        getOcrModelCatalog(),
        getOcrModelDownloadStatus(),
      ]);
      setModelCatalog(catalog);
      setModelDownload(downloadStatus);
    } catch {
      setModelCatalog(null);
    }
  }, []);

  const importModelPackage = useCallback(
    async (packagePath: string) => {
      if (!isTauriEnvironment() || modelBusy || modelDownloadActive) return;
      if (!isOcrModelPackagePath(packagePath)) {
        setError(
          isEn
            ? "Drop a VisualTeX .vtxocrmodel package here."
            : "Thả gói VisualTeX .vtxocrmodel vào đây.",
        );
        return;
      }

      const previouslyInstalled = new Set(runtime?.installedModels ?? []);
      const packageModel = modelNameFromPackagePath(packagePath);
      setModelBusy(true);
      setError("");
      try {
        const nextRuntime = await installOptionalOcrModel(packagePath);
        applyRuntimeStatus(nextRuntime);
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
    [isEn, model, modelBusy, modelDownloadActive, onModelChange, onNotify, runtime?.installedModels],
  );

  useEffect(() => {
    if (!open) return;
    setError("");

    let cancelled = false;
    const frame = window.requestAnimationFrame(() => {
      if (cancelled) return;
      // Re-read the pointer and actual runtime files every time the dialog is
      // opened. Keeping the previous React object here made a changed path or
      // deleted environment appear unchanged after reopening the dialog.
      void refreshRuntime(false);
      void refreshInstallStatus();
      void refreshModelCatalog();
    });
    return () => {
      cancelled = true;
      runtimeRequestGenerationRef.current += 1;
      window.cancelAnimationFrame(frame);
    };
  }, [open, refreshInstallStatus, refreshModelCatalog, refreshRuntime]);

  useEffect(() => {
    if (!open || !isTauriEnvironment() || !hasTauriWebviewRuntime()) return;
    let unlisten: (() => void) | undefined;
    let disposed = false;

    const updateDropHighlight = (paths: string[], position: { x: number; y: number }) => {
      const canImport = !modelBusy && !modelDownloadActive;
      const containsModelPackage = paths.some(isOcrModelPackagePath);
      setModelPackageDragging(
        canImport &&
          containsModelPackage &&
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
  }, [importModelPackage, isEn, modelBusy, modelDownloadActive, open]);

  useEffect(() => {
    if (!open || !isTauriEnvironment() || !hasTauriWebviewRuntime()) return;
    let unlisten: (() => void) | undefined;
    let cancelled = false;
    void listenOcrModelDownloadProgress((progress) => {
      if (cancelled) return;
      if (
        modelCancelRequestedRef.current &&
        progress.state !== "cancelled" &&
        progress.state !== "failed" &&
        progress.state !== "complete"
      ) {
        return;
      }
      setModelDownload(progress);
      if (progress.state === "complete") {
        modelCancelRequestedRef.current = false;
        setModelBusy(false);
        setError("");
        void refreshRuntime(true);
      } else if (progress.state === "failed" || progress.state === "cancelled") {
        modelCancelRequestedRef.current = false;
        setModelBusy(false);
      }
    }).then((dispose) => {
      if (cancelled) dispose();
      else unlisten = dispose;
    });
    return () => {
      cancelled = true;
      unlisten?.();
    };
  }, [open, refreshRuntime]);

  useEffect(() => {
    if (
      !open ||
      (!hasTauriWebviewRuntime() && !isOfficeCompanionEnvironment())
    ) {
      return;
    }
    let unlisten: (() => void) | undefined;
    let cancelled = false;
    void listenOcrInstallProgress((progress) => {
      if (cancelled) return;
      const active = isOcrInstallActive(progress.state);
      setInstallProgress(progress);
      installingRef.current = active;
      setInstalling(active);
      setInstallStatus((current) => ({
        schemaVersion: current?.schemaVersion ?? 1,
        state: progress.state,
        currentStep: progress.stage,
        completedSteps: current?.completedSteps ?? [],
        percent: progress.percent,
        message: progress.message,
        detail: progress.detail,
        error: progress.error,
        logPath: progress.logPath ?? current?.logPath ?? "",
        updatedAtMs: Date.now(),
      }));
      if (progress.state === "complete") {
        setError("");
        void refreshRuntime(false);
      }
    }).then((dispose) => {
      if (cancelled) dispose();
      else unlisten = dispose;
    });
    return () => {
      cancelled = true;
      unlisten?.();
    };
  }, [open, refreshRuntime]);

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
    if (!beginOcrInstallGuard(installingRef)) return;
    setInstalling(true);
    setError("");
    const startingProgress: OcrInstallProgress = {
      stage: installFailed ? installStatus?.currentStep ?? "resume" : "start",
      state: "installing",
      percent: installFailed ? installStatus?.percent ?? 1 : 1,
      message: installFailed
        ? isEn
          ? "Resuming from the failed OCR installation step"
          : "Tiếp tục từ bước cài đặt OCR không thành công"
        : isEn
          ? "Starting OCR installation"
          : "Bắt đầu cài đặt OCR",
      detail: isWindows
        ? isEn
          ? "Python 3.12 is preferred. Python 3.13 is incompatible with tokenizers 0.19.1 and will not be selected."
          : "Python 3.12 được ưu tiên. Python 3.13 không tương thích với mã thông báo 0.19.1 và sẽ không được chọn."
        : null,
      error: null,
      logPath: installStatus?.logPath ?? null,
    };
    setInstallProgress(startingProgress);

    try {
      const nextRuntime = await installOcrRuntime();
      applyRuntimeStatus(nextRuntime);
      setError("");
      await refreshInstallStatus();
      onNotify(isEn ? "OCR runtime installed" : "Đã cài đặt thời gian chạy OCR");
    } catch (installError) {
      const message = readError(installError);
      setError(message);
      await refreshInstallStatus();
    } finally {
      endOcrInstallGuard(installingRef);
      setInstalling(false);
    }
  };

  const handleCancelInstall = async () => {
    if (!installing) return;
    try {
      await cancelOcrInstall();
      onNotify(isEn ? "OCR installation cancellation requested" : "Đã yêu cầu hủy cài đặt OCR");
      await refreshInstallStatus();
    } catch (cancelError) {
      setError(readError(cancelError));
    }
  };

  const handleOpenInstallLogs = async () => {
    try {
      await openOcrInstallLogs();
    } catch (logError) {
      setError(readError(logError));
    }
  };

  const handleChangeStorage = async () => {
    if (
      !isTauriEnvironment() ||
      changingStorage ||
      installing ||
      modelBusy ||
      modelDownloadActive ||
      recognizing
    ) {
      return;
    }
    let reinstallGuardHeld = false;
    try {
      const selected = await openDialog({
        multiple: false,
        directory: true,
        title: isEn
          ? "Choose a parent folder for VisualTeX OCR storage"
          : "Chọn thư mục mẹ để lưu trữ VisualTeX OCR",
      });
      if (typeof selected !== "string") return;
      const hasExistingData =
        Boolean(runtime?.installed) ||
        installedModels.length > 0 ||
        Boolean(runtime?.pythonPath) ||
        (installStatus?.percent ?? 0) > 1;
      const confirmed = window.confirm(
        isEn
          ? hasExistingData
            ? `VisualTeX will create or use a VisualTeX-OCR folder under:\n${selected}\n\nThe current private Python environment, dependencies, models, resumable downloads, caches, and logs will be deleted. If the destination contains an incomplete VisualTeX OCR environment, that incomplete data will also be reset. VisualTeX will then switch to the new location and reinstall the OCR runtime there. Models must be downloaded or imported again. Continue?`
            : `VisualTeX will create or use a VisualTeX-OCR folder under:\n${selected}\n\nThe OCR runtime and all future models, downloads, caches, and logs will use this location. It will be preserved after uninstall. Continue?`
          : hasExistingData
            ? `VisualTeX will create or use the VisualTeX-OCR folder at: 
${selected}

The current private Python environment, all dependencies, models, breakpoint downloads, caches and logs will be deleted; if there is an incomplete VisualTeX OCR environment at the target location, it will also be safely reset. The OCR environment will then be switched to a new location and reinstalled, and the model will need to be downloaded or imported again. Continue?`
            : `VisualTeX will create or use the VisualTeX-OCR folder at: 
${selected}

The OCR environment as well as future installed models, downloads, caches and logs are written to this location and will remain after the software is uninstalled. Continue?`,
      );
      if (!confirmed) return;

      setChangingStorage(true);
      setError("");
      let nextRuntime = await configureOcrStorageLocation(selected);
      applyRuntimeStatus(nextRuntime);

      if (hasExistingData) {
        if (!beginOcrInstallGuard(installingRef)) {
          throw new Error(
            isEn ? "OCR installation is already running" : "Quá trình cài đặt OCR đang chạy",
          );
        }
        reinstallGuardHeld = true;
        setInstalling(true);
        setInstallProgress({
          stage: "start",
          state: "installing",
          percent: 1,
          message: isEn
            ? "Reinstalling OCR at the new storage location"
            : "Cài đặt lại OCR tại vị trí lưu trữ mới",
          detail: isEn
            ? "The previous environment was reset. Models are installed separately."
            : "Môi trường trước đó đã được đặt lại. Các mô hình được cài đặt riêng.",
          error: null,
          logPath: null,
        });
        nextRuntime = await installOcrRuntime();
        applyRuntimeStatus(nextRuntime);
      }

      const fallback = resolveAvailableOcrModel(nextRuntime, model);
      if (fallback !== model) onModelChange(fallback);
      await Promise.all([refreshInstallStatus(), refreshModelCatalog()]);
      onNotify(
        isEn
          ? hasExistingData
            ? "OCR storage changed and the runtime was reinstalled"
            : "OCR storage location updated"
          : hasExistingData
            ? "The OCR storage location has been changed and the operating environment has been reinstalled."
            : "OCR storage location updated",
      );
    } catch (storageError) {
      setError(readError(storageError));
      await refreshRuntime(true);
      await refreshInstallStatus();
    } finally {
      if (reinstallGuardHeld) endOcrInstallGuard(installingRef);
      setInstalling(false);
      setChangingStorage(false);
    }
  };

  const handleOpenStorage = async () => {
    try {
      await openOcrStorageLocation();
    } catch (storageError) {
      setError(readError(storageError));
    }
  };

  const handleImportModel = async () => {
    if (!isTauriEnvironment() || modelBusy || modelDownloadActive) return;
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

  const handleDownloadModel = async () => {
    if (!selectedCatalogEntry || modelBusy || modelDownloadActive) return;
    const confirmed = window.confirm(
      isEn
        ? `Download ${selectedModel.labelEn}? Expected download: ${readableBytes(selectedCatalogEntry.size)}. VisualTeX will keep a .part file for resume and verify SHA-256 before activation.`
        : `Tải xuống ${selectedModel.labelEn}? Dự kiến ​​tải xuống: ${readableBytes(selectedCatalogEntry.size)}. VisualTeX sẽ giữ tệp .part để tiếp tục và xác minh SHA-256 trước khi kích hoạt.`,
    );
    if (!confirmed) return;
    modelCancelRequestedRef.current = false;
    setModelBusy(true);
    setError("");
    try {
      const nextRuntime = await downloadOcrModel(model);
      applyRuntimeStatus(nextRuntime);
      onNotify(isEn ? "OCR model downloaded and verified" : "Đã tải xuống và xác minh mô hình OCR");
    } catch (downloadError) {
      const message = readError(downloadError);
      if (message.toLowerCase().includes("cancel")) setError("");
      else setError(message);
    } finally {
      setModelBusy(false);
    }
  };

  const handleCancelModelDownload = async () => {
    modelCancelRequestedRef.current = true;
    setModelDownload((current) =>
      current
        ? {
            ...current,
            state: "cancelled",
            speedBytesPerSecond: 0,
            etaSeconds: null,
            message: isEn
              ? "OCR model download cancelled immediately; the .part file was kept for resume"
              : "Việc tải xuống mô hình OCR bị hủy ngay lập tức; tập tin .part được giữ lại để làm sơ yếu lý lịch",
            error: null,
          }
        : current,
    );
    setError("");
    try {
      await cancelOcrModelDownload();
    } catch (cancelError) {
      setError(readError(cancelError));
    }
  };

  const handleRemoveModel = async () => {
    if (!selectedModelInstalled || modelBusy || modelDownloadActive) return;
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
      applyRuntimeStatus(nextRuntime);
      const fallback = resolveAvailableOcrModel(nextRuntime, DEFAULT_OCR_MODEL);
      onModelChange(fallback);
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
          ? "Import or explicitly download the selected OCR model first."
          : "Nhập hoặc tải xuống rõ ràng mô hình OCR đã chọn trước.",
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
      if (selectedModelInstalled) {
        void warmupOcrModel(model).catch((warmupError) => {
          setError(readError(warmupError));
        });
      }
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
      const nextRuntime = await resetOcrRuntime();
      applyRuntimeStatus(nextRuntime);
      await refreshInstallStatus();
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
                disabled={recognizing || cancelling || modelBusy || modelDownloadActive}
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
              {selectedModelInstalled ? <CheckCircle2 size={15} /> : <AlertCircle size={15} />}
              <div>
                <strong>
                  {selectedModelInstalled
                    ? isEn
                      ? `${selectedModel.labelEn} is installed`
                      : `${selectedModel.labelEn} đã được cài đặt`
                    : damagedModels.includes(model)
                      ? isEn
                        ? `${selectedModel.labelEn} was damaged and has been quarantined`
                        : `${selectedModel.labelEn} đã bị hỏng và đã bị cách ly`
                      : isEn
                        ? `${selectedModel.labelEn} is not installed`
                        : `${selectedModel.labelEn} chưa được cài đặt`}
                </strong>
                <span>
                  {selectedModelInstalled
                    ? isEn
                      ? "Recognition and warmup use only this verified local model_dir."
                      : "Việc nhận dạng và khởi động chỉ sử dụng model_dir cục bộ đã được xác minh này."
                    : isEn
                      ? "Import a verified .vtxocrmodel package, or explicitly confirm the catalog download."
                      : "Nhập gói .vtxocrmodel đã được xác minh hoặc xác nhận rõ ràng việc tải xuống danh mục."}
                </span>
                <small className="ocr-model-drop-hint">
                  {modelPackageDragging
                    ? isEn
                      ? "Release to verify and import this model package"
                      : "Phát hành để xác minh và nhập gói mô hình này"
                    : isEn
                      ? "You can also drag one .vtxocrmodel package directly into this area."
                      : "Bạn cũng có thể kéo trực tiếp một gói .vtxocrmodel vào khu vực này."}
                </small>
                {modelDownload?.model === model && (
                  <>
                    <div className="ocr-progress-label">
                      <span>{modelDownload.message}</span>
                      <strong>{modelDownload.percent}%</strong>
                    </div>
                    <div className="ocr-progress-track">
                      <span style={{ width: modelDownload.percent + "%" }} />
                    </div>
                    <small>
                      {readableBytes(modelDownload.downloadedBytes)} / {readableBytes(modelDownload.totalBytes)}
                      {modelDownload.speedBytesPerSecond > 0
                        ? ` · ${readableBytes(modelDownload.speedBytesPerSecond)}/s`
                        : ""}
                      {modelDownload.etaSeconds !== null
                        ? ` · ${isEn ? "ETA" : "ETA"} ${readableEta(modelDownload.etaSeconds)}`
                        : ""}
                    </small>
                    {modelDownload.error && <pre className="ocr-install-error">{modelDownload.error}</pre>}
                  </>
                )}
                <div className="ocr-install-actions">
                  {!selectedModelInstalled && (
                    <button
                      type="button"
                      onClick={() => void handleImportModel()}
                      disabled={!isTauriEnvironment() || modelBusy || modelDownloadActive}
                    >
                      <Upload size={14} />
                      {isEn ? "Import package" : "Gói hàng nhập khẩu"}
                    </button>
                  )}
                  {!selectedModelInstalled && selectedCatalogEntry && !modelDownloadActive && (
                    <button
                      type="button"
                      className="primary-button"
                      onClick={() => void handleDownloadModel()}
                      disabled={!runtime?.installed || modelBusy}
                    >
                      <Download size={14} />
                      {modelDownload?.state === "failed" || modelDownload?.state === "cancelled"
                        ? isEn
                          ? "Retry / resume"
                          : "Thử lại/tiếp tục"
                        : isEn
                          ? `Download ${readableBytes(selectedCatalogEntry.size)}`
                          : `Tải xuống ${readableBytes(selectedCatalogEntry.size)}`}
                    </button>
                  )}
                  {modelDownloadActive && modelDownload?.model === model && (
                    <button type="button" onClick={() => void handleCancelModelDownload()}>
                      <X size={14} />
                      {isEn ? "Cancel download" : "Hủy tải xuống"}
                    </button>
                  )}
                  {selectedModelInstalled && (
                    <button
                      type="button"
                      className="is-danger"
                      onClick={() => void handleRemoveModel()}
                      disabled={modelBusy || modelDownloadActive || recognizing}
                    >
                      <Trash2 size={14} />
                      {isEn ? "Remove model" : "Xóa mô hình"}
                    </button>
                  )}
                </div>
                {!selectedCatalogEntry && !selectedModelInstalled && isTauriEnvironment() && (
                  <small>
                    {isEn
                      ? "This build has no verified online catalog entry for the selected model; use manual import."
                      : "Bản dựng này không có mục nhập danh mục trực tuyến đã được xác minh cho kiểu máy đã chọn; sử dụng nhập thủ công."}
                  </small>
                )}
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
                <span className={"ocr-runtime-icon " + (runtime?.installed && !checkingRuntime && !changingStorage ? "is-ready" : "")}>
                  {checkingRuntime || changingStorage ? (
                    <LoaderCircle size={17} className="is-spinning" />
                  ) : runtime?.installed ? (
                    <CheckCircle2 size={17} />
                  ) : (
                    <Cpu size={17} />
                  )}
                </span>
                <div>
                  <strong>
                    {changingStorage
                      ? isEn
                        ? "Resetting and changing OCR storage"
                        : "Đặt lại và thay đổi bộ nhớ OCR"
                      : checkingRuntime
                        ? isEn
                          ? "Checking the actual OCR environment"
                          : "Kiểm tra môi trường OCR thực tế"
                        : runtime?.installed
                          ? isEn
                            ? "Local OCR runtime ready"
                            : "Thời gian chạy OCR cục bộ đã sẵn sàng"
                          : installing
                        ? isEn
                          ? "OCR runtime is being installed"
                          : "Thời gian chạy OCR đang được cài đặt"
                        : installStatus?.state === "verificationFailed"
                          ? isEn
                            ? "OCR runtime verification failed"
                            : "Xác minh thời gian chạy OCR không thành công"
                          : installFailed
                            ? isEn
                              ? "OCR installation failed"
                              : "Cài đặt OCR không thành công"
                            : isEn
                              ? "OCR runtime is not installed"
                              : "Thời gian chạy OCR chưa được cài đặt"}
                  </strong>
                  <span>
                    {changingStorage
                      ? isEn
                        ? "The displayed path and installation state will update after the disk operation completes."
                        : "Đường dẫn hiển thị và trạng thái cài đặt sẽ cập nhật sau khi thao tác trên đĩa hoàn tất."
                      : installing || installFailed
                        ? installStatus?.message ?? installProgress?.message
                        : runtime?.message ?? (isEn ? "Checking runtime…" : "Đang kiểm tra thời gian chạy…")}
                  </span>
                </div>
              </div>

              {runtime && (
                <div
                  className={
                    "ocr-storage-location" +
                    (storageLowForInitialInstall ? " is-low-space" : "")
                  }
                >
                  <div className="ocr-storage-location-main">
                    <HardDrive size={16} />
                    <div>
                      <span>{isEn ? "Independent OCR storage" : "Lưu trữ OCR độc lập"}</span>
                      <code title={runtime.runtimePath || undefined}>
                        {runtime.runtimePath || (isEn ? "Unavailable" : "Không có")}
                      </code>
                    </div>
                  </div>
                  <div className="ocr-storage-location-meta">
                    <span>
                      {storageAvailableBytes === null
                        ? isEn
                          ? "Free space unavailable"
                          : "Không có dung lượng trống"
                        : isEn
                          ? `${readableBytes(storageAvailableBytes)} free`
                          : `${readableBytes(storageAvailableBytes)} miễn phí`}
                    </span>
                    <span>
                      {runtime.storageSource === "legacy"
                        ? isEn
                          ? "Existing environment adopted"
                          : "Đã áp dụng môi trường hiện tại"
                        : runtime.storagePersistentAcrossUninstall
                          ? isEn
                            ? "Preserved after uninstall"
                            : "Bảo quản sau khi gỡ cài đặt"
                          : isEn
                            ? "Application data location"
                            : "Vị trí dữ liệu ứng dụng"}
                    </span>
                  </div>
                  {storageLowForInitialInstall && (
                    <small className="ocr-storage-space-warning">
                      {isEn
                        ? "Less than 2 GB is available. Choose another disk before installing the OCR runtime."
                        : "Có sẵn dưới 2 GB. Chọn một đĩa khác trước khi cài đặt thời gian chạy OCR."}
                    </small>
                  )}
                  {isTauriEnvironment() && (
                    <div className="ocr-storage-location-actions">
                      <button
                        type="button"
                        onClick={() => void handleChangeStorage()}
                        disabled={
                          changingStorage ||
                          installing ||
                          modelBusy ||
                          modelDownloadActive ||
                          recognizing
                        }
                      >
                        {changingStorage ? (
                          <LoaderCircle size={13} className="is-spinning" />
                        ) : (
                          <FolderOpen size={13} />
                        )}
                        {changingStorage
                          ? isEn
                            ? "Resetting and switching…"
                            : "Đặt lại và chuyển đổi…"
                          : isEn
                            ? "Change location"
                            : "Thay đổi địa điểm"}
                      </button>
                      <button
                        type="button"
                        onClick={() => void handleOpenStorage()}
                        disabled={changingStorage || !runtime.runtimePath}
                      >
                        <FolderOpen size={13} />
                        {isEn ? "Open folder" : "Mở thư mục"}
                      </button>
                    </div>
                  )}
                </div>
              )}

              {runtime?.installed ? (
                <div className="ocr-runtime-details">
                  <span>Python {runtime.pythonVersion}</span>
                  <span>Paddle {runtime.paddleVersion}</span>
                  <span>PaddleOCR {runtime.paddleocrVersion}</span>
                  <span>
                    {isEn ? "Models" : "Model"}: {installedModels.length > 0 ? installedModels.join(", ") : isEn ? "none" : "không có"}
                  </span>
                  <button type="button" onClick={handleRestartWorker} disabled={!selectedModelInstalled}>
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
                      <div className="ocr-install-actions">
                        <button type="button" onClick={() => void handleCancelInstall()}>
                          <X size={14} />
                          {isEn ? "Cancel installation" : "Hủy cài đặt"}
                        </button>
                        <button type="button" onClick={() => void handleOpenInstallLogs()}>
                          <ScanLine size={14} />
                          {isEn ? "View log" : "Xem nhật ký"}
                        </button>
                      </div>
                    </>
                  ) : installFailed ? (
                    <>
                      <div className="ocr-progress-label is-failed">
                        <span>{installStatus?.message ?? (isEn ? "Installation failed" : "Cài đặt không thành công")}</span>
                        <strong>{installStatus?.percent ?? installProgress?.percent ?? 0}%</strong>
                      </div>
                      <div className="ocr-progress-track is-failed">
                        <span
                          style={{
                            width:
                              (installStatus?.percent ?? installProgress?.percent ?? 0) + "%",
                          }}
                        />
                      </div>
                      {(installStatus?.detail ?? installProgress?.detail) && (
                        <small>{installStatus?.detail ?? installProgress?.detail}</small>
                      )}
                      {(installStatus?.error ?? installProgress?.error) && (
                        <pre className="ocr-install-error">
                          {installStatus?.error ?? installProgress?.error}
                        </pre>
                      )}
                      <div className="ocr-install-actions">
                        <button
                          type="button"
                          className="primary-button"
                          onClick={() => void handleInstall()}
                          disabled={changingStorage || storageLowForInitialInstall}
                        >
                          <RefreshCw size={14} />
                          {isEn ? "Retry current step" : "Thử lại bước hiện tại"}
                        </button>
                        <button type="button" onClick={() => void handleOpenInstallLogs()}>
                          <ScanLine size={14} />
                          {isEn ? "View log" : "Xem nhật ký"}
                        </button>
                        <button
                          type="button"
                          className="is-danger"
                          onClick={() => void handleResetRuntime()}
                        >
                          <Trash2 size={14} />
                          {isEn ? "Reset environment" : "Đặt lại môi trường"}
                        </button>
                      </div>
                    </>
                  ) : (
                    <>
                      <p>
                        {isWindows
                          ? isEn
                            ? "VisualTeX installs the bundled private Python 3.12.10 x64 and fixed wheelhouse with --no-index and --find-links. PyPI, system Python, and user site-packages are never used. Models are managed separately."
                            : "VisualTeX cài đặt gói Python 3.12.10 x64 riêng tư và bộ điều khiển cố định với --no-index và --find-links. PyPI, Python hệ thống và gói trang web của người dùng không bao giờ được sử dụng. Các mô hình được quản lý riêng biệt."
                          : isEn
                            ? "VisualTeX verifies and installs the bundled offline OCR runtime. Recognition models are managed separately."
                            : "VisualTeX xác minh và cài đặt thời gian chạy OCR ngoại tuyến đi kèm. Các mô hình nhận dạng được quản lý riêng biệt."}
                      </p>
                      <div className="ocr-install-actions">
                        <button
                          type="button"
                          className="primary-button"
                          onClick={() => void handleInstall()}
                          disabled={
                            (!isTauriEnvironment() &&
                              !isOfficeCompanionEnvironment()) ||
                            checkingRuntime ||
                            installing ||
                            changingStorage ||
                            storageLowForInitialInstall
                          }
                        >
                          <Download size={15} />
                          {isEn ? "Install OCR runtime" : "Cài đặt thời gian chạy OCR"}
                        </button>
                        <button type="button" onClick={() => void handleOpenInstallLogs()}>
                          <ScanLine size={14} />
                          {isEn ? "View log" : "Xem nhật ký"}
                        </button>
                      </div>
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
                      ? "Recognition uses only the verified local model_dir. It will never download a model in the background."
                      : "Việc nhận dạng chỉ sử dụng model_dir cục bộ đã được xác minh. Nó sẽ không bao giờ tải xuống một mô hình ở chế độ nền."}
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

            {shouldDisplayRuntimeError(error, installStatus?.state) && (
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
                modelBusy ||
                modelDownloadActive
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
