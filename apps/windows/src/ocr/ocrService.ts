export type UnlistenFn = () => void;

export interface OcrTransportEvent<T> {
  event: string;
  id: number;
  payload: T;
}

export interface OcrTransport {
  environment: "desktop" | "office";
  invoke<T>(command: string, args?: Record<string, unknown>): Promise<T>;
  listen<T>(
    eventName: string,
    handler: (event: OcrTransportEvent<T>) => void,
  ): Promise<UnlistenFn>;
}

let configuredTransport: OcrTransport | null = null;

export function configureOcrTransport(transport: OcrTransport) {
  configuredTransport = transport;
}

function activeTransport() {
  if (!configuredTransport) {
    throw new Error("VisualTeX OCR transport has not been initialized.");
  }
  return configuredTransport;
}

function invoke<T>(command: string, args?: Record<string, unknown>) {
  return activeTransport().invoke<T>(command, args);
}

function listen<T>(
  eventName: string,
  handler: (event: OcrTransportEvent<T>) => void,
) {
  return activeTransport().listen(eventName, handler);
}

export const OCR_MODELS = [
  {
    id: "PP-FormulaNet_plus-S",
    labelVi: "Nhanh S",
    labelEn: "Fast S",
    hintVi: "Gói mô hình riêng biệt, cài đặt khoảng 248 MB; nhanh nhất cho công thức tiếng Anh",
    hintEn: "Separate model package, about 248 MB installed; fastest for English formulas",
    downloadMb: 259.6,
    storageMb: 248,
    cpuBenchmarkMs: 260.99,
  },
  {
    id: "PP-FormulaNet_plus-M",
    labelVi: "Cân bằng M (khuyên dùng)",
    labelEn: "Balanced M (recommended)",
    hintVi: "Gói mô hình riêng biệt; cân bằng cho các công thức phức tạp, được khuyên dùng",
    hintEn: "Separate model package; balanced for complex formulas, recommended",
    downloadMb: 620.5,
    storageMb: 592,
    cpuBenchmarkMs: 1615.8,
  },
  {
    id: "PP-FormulaNet_plus-L",
    labelVi: "Độ chính xác cao L",
    labelEn: "High accuracy L",
    hintVi: "Gói mô hình riêng biệt, cài đặt khoảng 698 MB; lần tải đầu tiên chậm và có thể sử dụng vài GB bộ nhớ",
    hintEn: "Separate model package, about 698 MB installed; first load is slow and may use several GB of memory",
    downloadMb: 731.5,
    storageMb: 698,
    cpuBenchmarkMs: 3125.58,
  },
] as const;

export type OcrModelName = (typeof OCR_MODELS)[number]["id"];
export const DEFAULT_OCR_MODEL: OcrModelName = "PP-FormulaNet_plus-M";

export interface OcrRuntimeStatus {
  installed: boolean;
  pythonPath: string | null;
  pythonVersion: string | null;
  paddleVersion: string | null;
  paddleocrVersion: string | null;
  runtimePath: string;
  storageConfigPath: string;
  storageSource: "configured" | "legacy" | "default" | string;
  storageManaged: boolean;
  storageAvailableBytes: number | null;
  storagePersistentAcrossUninstall: boolean;
  runtimeBundleAvailable: boolean;
  offlineBundleAvailable: boolean;
  installedModels: string[];
  damagedModels: string[];
  modelCatalogAvailable: boolean;
  defaultModel: string;
  message: string;
}

export function resolveAvailableOcrModel(
  runtime: Pick<OcrRuntimeStatus, "installedModels" | "defaultModel">,
  requested: OcrModelName,
): OcrModelName {
  const installed = new Set(runtime.installedModels);
  if (installed.has(requested)) return requested;
  if (installed.has(runtime.defaultModel)) {
    return runtime.defaultModel as OcrModelName;
  }
  const fallback = OCR_MODELS.find((item) => installed.has(item.id));
  return fallback?.id ?? requested;
}

export type OcrInstallState =
  | "notInstalled"
  | "installing"
  | "installFailed"
  | "dependenciesInstalled"
  | "verifying"
  | "verificationFailed"
  | "complete"
  | "cancelled";

export interface OcrInstallProgress {
  stage: string;
  state: OcrInstallState;
  percent: number;
  message: string;
  detail: string | null;
  error: string | null;
  logPath: string | null;
}

export interface OcrInstallStatus {
  schemaVersion: number;
  state: OcrInstallState;
  currentStep: string | null;
  completedSteps: string[];
  percent: number;
  message: string;
  detail: string | null;
  error: string | null;
  logPath: string;
  updatedAtMs: number;
}

export interface OcrModelCatalogEntry {
  model: OcrModelName;
  url: string;
  size: number;
  sha256: string;
}

export interface OcrModelCatalog {
  schemaVersion: number;
  platform: "windows";
  architecture: "x64";
  entries: OcrModelCatalogEntry[];
}

export type OcrModelDownloadState =
  | "idle"
  | "downloading"
  | "verifying"
  | "installing"
  | "complete"
  | "cancelled"
  | "failed";

export interface OcrModelDownloadSnapshot {
  model: OcrModelName;
  state: OcrModelDownloadState;
  downloadedBytes: number;
  totalBytes: number;
  percent: number;
  speedBytesPerSecond: number;
  etaSeconds: number | null;
  message: string;
  error: string | null;
}

export interface OcrFormulaResult {
  latex: string;
}

export interface OcrRecognitionProgress {
  event: "progress";
  id: string;
  stage: "preprocess" | "model" | "inference" | string;
  message: string;
  model: OcrModelName;
}

export interface OcrRecognitionResult {
  model: string;
  elapsedMs: number;
  processedWidth: number;
  processedHeight: number;
  backgroundInverted: boolean;
  backgroundLuminance: number;
  formulas: OcrFormulaResult[];
}

export interface OcrImageRequest {
  bytes: number[];
  extension: string;
  model: OcrModelName;
}

const SUPPORTED_EXTENSIONS = new Set([
  "png",
  "jpg",
  "jpeg",
  "webp",
  "bmp",
  "tif",
  "tiff",
]);

export const isTauriEnvironment = () =>
  configuredTransport?.environment === "desktop";

export const isOfficeCompanionEnvironment = () => {
  if (typeof window === "undefined" || typeof document === "undefined") {
    return false;
  }
  const token =
    window.__VISUALTEX_INSTALL_TOKEN__ ??
    document
      .querySelector<HTMLMetaElement>('meta[name="visualtex-install-token"]')
      ?.content;
  return (
    configuredTransport?.environment === "office" &&
    window.location.protocol === "https:" &&
    window.location.hostname === "127.0.0.1" &&
    window.location.port === "43127" &&
    typeof token === "string" &&
    token.length >= 32
  );
};

export function getImageExtension(file: File): string {
  const fromName = file.name.split(".").pop()?.toLocaleLowerCase() ?? "";
  if (SUPPORTED_EXTENSIONS.has(fromName)) return fromName;

  const mimeMap: Record<string, string> = {
    "image/png": "png",
    "image/jpeg": "jpg",
    "image/webp": "webp",
    "image/bmp": "bmp",
    "image/tiff": "tiff",
  };
  const fromMime = mimeMap[file.type];
  if (fromMime) return fromMime;
  throw new Error("This image format is not supported, please use PNG, JPEG, WebP, BMP or TIFF");
}

export function validateOcrImage(file: File) {
  getImageExtension(file);
  if (file.size <= 0) throw new Error("The picture file is empty");
  if (file.size > 20 * 1024 * 1024) {
    throw new Error("Image cannot exceed 20 MB");
  }
}

export async function fileToOcrRequest(
  file: File,
  model: OcrModelName,
): Promise<OcrImageRequest> {
  validateOcrImage(file);
  const buffer = await file.arrayBuffer();
  return {
    bytes: Array.from(new Uint8Array(buffer)),
    extension: getImageExtension(file),
    model,
  };
}

function requireOcrEnvironment() {
  if (!isTauriEnvironment() && !isOfficeCompanionEnvironment()) {
    throw new Error(
      "OCR is only available in the VisualTeX desktop app or the local Office editor.",
    );
  }
}

function requireDesktopOcrEnvironment() {
  if (!isTauriEnvironment()) {
    throw new Error("The optional OCR model package can only be managed in the VisualTeX desktop app.");
  }
}

export async function getOcrRuntimeStatus(
  forceRefresh = false,
): Promise<OcrRuntimeStatus> {
  requireOcrEnvironment();
  return invoke<OcrRuntimeStatus>("get_ocr_runtime_status", { forceRefresh });
}

export async function configureOcrStorageLocation(
  selectedDirectory: string,
): Promise<OcrRuntimeStatus> {
  requireDesktopOcrEnvironment();
  return invoke<OcrRuntimeStatus>("configure_ocr_storage_location", {
    selectedDirectory,
  });
}

export async function openOcrStorageLocation(): Promise<void> {
  requireDesktopOcrEnvironment();
  return invoke("open_ocr_storage_location");
}

export async function installOcrRuntime(): Promise<OcrRuntimeStatus> {
  requireOcrEnvironment();
  return invoke<OcrRuntimeStatus>("install_ocr_runtime");
}

export async function getOcrInstallStatus(): Promise<OcrInstallStatus> {
  requireOcrEnvironment();
  return invoke<OcrInstallStatus>("get_ocr_install_status");
}

export async function cancelOcrInstall(): Promise<void> {
  requireOcrEnvironment();
  return invoke("cancel_ocr_install");
}

export async function openOcrInstallLogs(): Promise<void> {
  requireOcrEnvironment();
  return invoke("open_ocr_install_logs");
}

export async function recognizeFormulaImage(
  request: OcrImageRequest,
): Promise<OcrRecognitionResult> {
  requireOcrEnvironment();
  return invoke<OcrRecognitionResult>("recognize_formula_image", { request });
}

export async function cancelOcrRecognition(): Promise<void> {
  requireOcrEnvironment();
  return invoke("cancel_ocr_recognition");
}

export async function restartOcrWorker(): Promise<void> {
  requireOcrEnvironment();
  return invoke("restart_ocr_worker");
}

export async function warmupOcrModel(model: OcrModelName): Promise<void> {
  requireOcrEnvironment();
  return invoke("warmup_ocr_model", { model });
}

export async function resetOcrRuntime(): Promise<OcrRuntimeStatus> {
  requireOcrEnvironment();
  return invoke<OcrRuntimeStatus>("reset_ocr_runtime");
}

export async function installOptionalOcrModel(
  packagePath: string,
): Promise<OcrRuntimeStatus> {
  requireDesktopOcrEnvironment();
  return invoke<OcrRuntimeStatus>("install_optional_ocr_model", {
    packagePath,
  });
}

export async function removeOptionalOcrModel(
  model: OcrModelName,
): Promise<OcrRuntimeStatus> {
  requireDesktopOcrEnvironment();
  return invoke<OcrRuntimeStatus>("remove_optional_ocr_model", { model });
}

export async function getOcrModelCatalog(): Promise<OcrModelCatalog> {
  requireDesktopOcrEnvironment();
  return invoke<OcrModelCatalog>("get_ocr_model_catalog");
}

export async function getOcrModelDownloadStatus(): Promise<OcrModelDownloadSnapshot | null> {
  requireDesktopOcrEnvironment();
  return invoke<OcrModelDownloadSnapshot | null>("get_ocr_model_download_status");
}

export async function downloadOcrModel(
  model: OcrModelName,
): Promise<OcrRuntimeStatus> {
  requireDesktopOcrEnvironment();
  return invoke<OcrRuntimeStatus>("download_ocr_model", { model });
}

export async function cancelOcrModelDownload(): Promise<boolean> {
  requireDesktopOcrEnvironment();
  return invoke<boolean>("cancel_ocr_model_download");
}

export async function listenOcrModelDownloadProgress(
  listener: (progress: OcrModelDownloadSnapshot) => void,
): Promise<UnlistenFn> {
  requireDesktopOcrEnvironment();
  return listen<OcrModelDownloadSnapshot>("ocr-model-download-progress", (event) => {
    listener(event.payload);
  });
}

export async function listenOcrRecognitionProgress(
  listener: (progress: OcrRecognitionProgress) => void,
): Promise<UnlistenFn> {
  requireOcrEnvironment();
  return listen<OcrRecognitionProgress>("ocr-recognition-progress", (event) => {
    listener(event.payload);
  });
}

export async function listenOcrInstallProgress(
  listener: (progress: OcrInstallProgress) => void,
): Promise<UnlistenFn> {
  requireOcrEnvironment();
  return listen<OcrInstallProgress>("ocr-install-progress", (event) => {
    listener(event.payload);
  });
}
