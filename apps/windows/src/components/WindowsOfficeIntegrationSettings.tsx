import { useCallback, useEffect, useMemo, useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import {
  CheckCircle2,
  ChevronDown,
  CircleAlert,
  CircleHelp,
  Download,
  ExternalLink,
  FolderOpen,
  Play,
  RefreshCw,
  Settings2,
  ShieldAlert,
  Square,
  ToggleLeft,
  ToggleRight,
  Trash2,
  Wrench,
  X,
} from "lucide-react";
import { useEditorStore } from "../stores/editorStore";

export type WindowsOfficeMode = "auto" | "vsto";

interface OfficePlatformStatus {
  platform: string;
  mode: WindowsOfficeMode;
  activeBackend: string;
  oleBridgeHealthy: boolean;
  oleLocalServerHealthy: boolean;
  staticInstallVerified: boolean;
  wordFilesPresent: boolean;
  wordRegistryComplete: boolean;
  wordLoadEnabled: boolean;
  powerpointFilesPresent: boolean;
  powerpointRegistryComplete: boolean;
  powerpointLoadEnabled: boolean;
  vstoWordHealthy: boolean;
  vstoPowerpointHealthy: boolean;
  wordConnected: boolean;
  powerpointConnected: boolean;
  connectionVerificationAttempted: boolean;
  companionProcessRunning: boolean;
  companionPortListening: boolean;
  companionHttpsHealthy: boolean;
  companionCertificateMatches: boolean;
  companionProtocolMatches: boolean;
  officeRuntimeVerified: boolean;
  currentUserCertificateTrusted: boolean;
  backgroundStartEnabled: boolean;
  lastError: string | null;
}

interface OfficeCompanionStatus {
  running: boolean;
  bindAddress: string;
  port: number;
  certificatePath: string;
  officeUiVersion: string;
  protocolVersion: number;
  lastError: string | null;
}

function errorMessage(error: unknown, fallback: string) {
  if (error instanceof Error && error.message) return error.message;
  if (typeof error === "string" && error.trim()) return error;
  return fallback;
}

function StatusLine({ ok, children }: { ok: boolean; children: React.ReactNode }) {
  return (
    <div className="office-platform-status-line">
      {ok ? (
        <CheckCircle2 className="office-state-ok" size={15} />
      ) : (
        <ShieldAlert className="office-state-warning" size={15} />
      )}
      <span>{children}</span>
    </div>
  );
}

export function WindowsOfficeIntegrationSettings() {
  const isEn = useEditorStore((state) => state.language) === "en";
  const [status, setStatus] = useState<OfficePlatformStatus | null>(null);
  const [companion, setCompanion] = useState<OfficeCompanionStatus | null>(null);
  const [busy, setBusy] = useState<string | null>(null);
  const [message, setMessage] = useState("");
  const [confirmRuntimeTest, setConfirmRuntimeTest] = useState(false);
  const [forceCloseOffice, setForceCloseOffice] = useState(false);
  const [confirmUninstall, setConfirmUninstall] = useState(false);

  const refresh = useCallback(async () => {
    setBusy((value) => value ?? "refresh");
    try {
      const [nextStatus, nextCompanion] = await Promise.all([
        invoke<OfficePlatformStatus>("get_office_platform_status"),
        invoke<OfficeCompanionStatus>("get_office_companion_status"),
      ]);
      setStatus(nextStatus);
      setCompanion(nextCompanion);
      setMessage("");
    } catch (error) {
      setMessage(
        errorMessage(
          error,
          isEn
            ? "Unable to read Windows Office integration status."
            : "Không thể đọc trạng thái tích hợp Windows Office.",
        ),
      );
    } finally {
      setBusy((value) => (value === "refresh" ? null : value));
    }
  }, [isEn]);

  useEffect(() => {
    void refresh();
  }, [refresh]);

  const run = useCallback(
    async (name: string, command: string, args?: Record<string, unknown>) => {
      setBusy(name);
      setMessage("");
      try {
        await invoke(command, args);
        if (
          command !== "open_word" &&
          command !== "open_powerpoint" &&
          command !== "open_windows_office_logs"
        ) {
          await refresh();
        }
        return true;
      } catch (error) {
        setMessage(
          errorMessage(
            error,
            isEn ? "Windows Office operation failed." : "Thao tác Windows Office không thành công.",
          ),
        );
        return false;
      } finally {
        setBusy(null);
      }
    },
    [isEn, refresh],
  );

  const installHealthy = Boolean(
    status?.staticInstallVerified &&
      status.wordFilesPresent &&
      status.wordRegistryComplete &&
      status.wordLoadEnabled &&
      status.powerpointFilesPresent &&
      status.powerpointRegistryComplete &&
      status.powerpointLoadEnabled &&
      status.oleLocalServerHealthy,
  );
  const runtimeHealthy = Boolean(status?.officeRuntimeVerified);
  const officeConnectionsVerified = Boolean(
    status?.wordConnected && status?.powerpointConnected,
  );
  const integrationReady = Boolean(
    installHealthy &&
      runtimeHealthy &&
      officeConnectionsVerified &&
      !status?.lastError,
  );
  const hasInstalledComponents = Boolean(
    status?.wordFilesPresent ||
      status?.powerpointFilesPresent ||
      status?.wordRegistryComplete ||
      status?.powerpointRegistryComplete ||
      status?.oleLocalServerHealthy,
  );
  const verificationPending = Boolean(
    installHealthy &&
      runtimeHealthy &&
      !officeConnectionsVerified &&
      !status?.connectionVerificationAttempted &&
      !status?.lastError,
  );
  const installationNeedsRepair = hasInstalledComponents && !installHealthy;
  const statusCopy = useMemo(() => {
    if (!status) {
      return {
        title: isEn ? "Checking Office integration…" : "Đang kiểm tra tích hợp Office…",
        description: isEn
          ? "VisualTeX is reading the Word and PowerPoint add-in status."
          : "VisualTeX đang đọc trạng thái bổ trợ Word và PowerPoint.",
      };
    }
    if (integrationReady) {
      return {
        title: isEn ? "Office integration is ready" : "Tích hợp Office đã sẵn sàng",
        description: isEn
          ? "Word and PowerPoint can create and edit VisualTeX formulas."
          : "Word và PowerPoint có thể tạo và chỉnh sửa công thức VisualTeX.",
      };
    }
    if (verificationPending) {
      return {
        title: isEn
          ? "Office integration is installed"
          : "Đã cài đặt tích hợp Office",
        description: isEn
          ? "Start Word and PowerPoint once to verify that both add-ins connect successfully."
          : "Khởi động Word và PowerPoint một lần để xác minh rằng cả hai phần bổ trợ đều kết nối thành công.",
      };
    }
    if (hasInstalledComponents) {
      return {
        title: installationNeedsRepair
          ? isEn
            ? "Office integration needs repair"
            : "Tích hợp văn phòng cần sửa chữa"
          : isEn
            ? "Office connection verification needs attention"
            : "Việc xác minh kết nối văn phòng cần được chú ý",
        description: installationNeedsRepair
          ? isEn
            ? "Some installation components are incomplete. Repair the integration and try again."
            : "Một số thành phần cài đặt chưa đầy đủ. Sửa chữa tích hợp và thử lại."
          : isEn
            ? "Close Office and run the connection verification again."
            : "Đóng Office và chạy lại xác minh kết nối.",
      };
    }
    return {
      title: isEn ? "Office integration is not installed" : "Chưa cài đặt tích hợp Office",
      description: isEn
        ? "Install once to add VisualTeX tools to Word and PowerPoint."
        : "Cài đặt một lần để thêm công cụ VisualTeX vào Word và PowerPoint.",
    };
  }, [
    hasInstalledComponents,
    installationNeedsRepair,
    integrationReady,
    isEn,
    status,
    verificationPending,
  ]);

  const diagnosticMessage =
    message ||
    (!verificationPending ? status?.lastError : null) ||
    companion?.lastError;

  const openRuntimeVerification = () => {
    setMessage("");
    setForceCloseOffice(false);
    setConfirmRuntimeTest(true);
  };

  const verifyRuntime = async () => {
    const succeeded = await run(
      "runtime-test",
      "test_windows_office_runtime",
      { forceCloseOffice },
    );
    if (succeeded) {
      setConfirmRuntimeTest(false);
      setForceCloseOffice(false);
    }
  };

  const uninstall = async () => {
    const succeeded = await run(
      "uninstall-ole",
      "uninstall_windows_ole_integration",
    );
    if (succeeded) setConfirmUninstall(false);
  };

  return (
    <section className="settings-section office-integration-section">
      <div className="office-settings-heading office-settings-heading-simple">
        <div>
          <h3>{isEn ? "Office integration" : "Tích hợp văn phòng"}</h3>
          <p>
            {isEn
              ? "Use VisualTeX directly in Microsoft Word and PowerPoint."
              : "Sử dụng VisualTeX trực tiếp trong Microsoft Word và PowerPoint."}
          </p>
        </div>
        <button
          type="button"
          className="icon-button compact"
          onClick={() => void refresh()}
          disabled={busy !== null}
          aria-label={isEn ? "Refresh Office status" : "Làm mới trạng thái Office"}
          title={isEn ? "Refresh status" : "Trạng thái làm mới"}
        >
          <RefreshCw
            size={15}
            className={busy === "refresh" ? "is-spinning" : ""}
          />
        </button>
      </div>

      <div
        className={`office-summary-card ${
          integrationReady
            ? "is-ready"
            : verificationPending
              ? "is-pending"
              : hasInstalledComponents
                ? "needs-attention"
                : "is-not-installed"
        }`}
      >
        <span className="office-summary-icon" aria-hidden="true">
          {integrationReady ? (
            <CheckCircle2 size={22} />
          ) : verificationPending ? (
            <CircleHelp size={22} />
          ) : (
            <CircleAlert size={22} />
          )}
        </span>
        <div>
          <strong>{statusCopy.title}</strong>
          <p>{statusCopy.description}</p>
        </div>
      </div>

      {diagnosticMessage && (
        <div className="office-settings-warning" role="alert">
          <ShieldAlert size={15} />
          <span className="office-settings-diagnostic">
            {diagnosticMessage}
          </span>
        </div>
      )}

      <div className="office-primary-actions">
        {!integrationReady && !hasInstalledComponents && (
          <button
            type="button"
            className="primary-button office-action-main"
            disabled={busy !== null}
            onClick={() =>
              void run("install-ole", "install_windows_ole_integration")
            }
          >
            <Download size={16} />
            {busy === "install-ole"
              ? isEn
                ? "Installing…"
                : "Đang cài đặt…"
              : isEn
                ? "Install Office integration"
                : "Cài đặt tích hợp Office"}
          </button>
        )}
        {installationNeedsRepair && (
          <button
            type="button"
            className="primary-button office-action-main"
            disabled={busy !== null}
            onClick={() =>
              void run("repair", "repair_windows_office_integration")
            }
          >
            <Wrench size={16} />
            {busy === "repair"
              ? isEn
                ? "Repairing…"
                : "Sửa chữa…"
              : isEn
                ? "Repair integration"
                : "Tích hợp sửa chữa"}
          </button>
        )}
        {hasInstalledComponents && installHealthy && !integrationReady && (
          <button
            type="button"
            className="primary-button office-action-main"
            disabled={busy !== null}
            onClick={openRuntimeVerification}
          >
            <CheckCircle2 size={16} />
            {busy === "runtime-test"
              ? isEn
                ? "Verifying…"
                : "Đang xác minh…"
              : status?.connectionVerificationAttempted
                ? isEn
                  ? "Verify Office connection again"
                  : "Xác minh lại kết nối Office"
                : isEn
                  ? "Verify Office connection"
                  : "Xác minh kết nối Office"}
          </button>
        )}
        <button
          type="button"
          className="secondary-button"
          disabled={busy !== null}
          onClick={() => void run("word", "open_word")}
        >
          <ExternalLink size={15} />
          {isEn ? "Open Word" : "Lời mở"}
        </button>
        <button
          type="button"
          className="secondary-button"
          disabled={busy !== null}
          onClick={() => void run("powerpoint", "open_powerpoint")}
        >
          <ExternalLink size={15} />
          {isEn ? "Open PowerPoint" : "Mở PowerPoint"}
        </button>
      </div>

      <details className="office-advanced-settings">
        <summary>
          <span>
            <Settings2 size={15} />
            {isEn ? "Advanced diagnostics" : "Chẩn đoán nâng cao"}
          </span>
          <ChevronDown size={15} className="office-details-chevron" />
        </summary>

        <div className="office-advanced-content">
          <div className="office-status-grid office-status-grid-compact">
            <article className="office-status-card">
              <header>
                <strong>{isEn ? "Installation" : "Cài đặt"}</strong>
                <StatusLine ok={installHealthy}>
                  {installHealthy
                    ? isEn
                      ? "Complete"
                      : "Hoàn thành"
                    : isEn
                      ? "Incomplete"
                      : "Chưa hoàn thành"}
                </StatusLine>
              </header>
              <StatusLine ok={Boolean(status?.wordFilesPresent)}>
                {isEn ? "Word add-in files" : "Tệp bổ trợ Word"}
              </StatusLine>
              <StatusLine ok={Boolean(status?.wordRegistryComplete)}>
                {isEn ? "Word registration" : "Đăng ký từ"}
              </StatusLine>
              <StatusLine ok={Boolean(status?.wordLoadEnabled)}>
                {isEn ? "Word add-in enabled" : "Đã bật phần bổ trợ Word"}
              </StatusLine>
              <StatusLine ok={Boolean(status?.powerpointFilesPresent)}>
                {isEn ? "PowerPoint add-in files" : "Tệp bổ trợ PowerPoint"}
              </StatusLine>
              <StatusLine ok={Boolean(status?.powerpointRegistryComplete)}>
                {isEn ? "PowerPoint registration" : "Đăng ký PowerPoint"}
              </StatusLine>
              <StatusLine ok={Boolean(status?.powerpointLoadEnabled)}>
                {isEn ? "PowerPoint add-in enabled" : "Đã bật bổ trợ PowerPoint"}
              </StatusLine>
              <StatusLine ok={Boolean(status?.oleLocalServerHealthy)}>
                {isEn ? "Formula OLE service" : "Dịch vụ công thức OLE"}
              </StatusLine>
            </article>

            <article className="office-status-card">
              <header>
                <strong>{isEn ? "Runtime" : "Thời gian chạy"}</strong>
                <StatusLine ok={runtimeHealthy}>
                  {runtimeHealthy
                    ? isEn
                      ? "Available"
                      : "Có sẵn"
                    : isEn
                      ? "Unavailable"
                      : "Không có"}
                </StatusLine>
              </header>
              <StatusLine ok={Boolean(status?.companionProcessRunning)}>
                {isEn ? "Companion process" : "Quá trình đồng hành"}
              </StatusLine>
              <StatusLine ok={Boolean(status?.companionPortListening)}>
                {isEn ? "Local port" : "Cảng địa phương"}
              </StatusLine>
              <StatusLine ok={Boolean(status?.companionHttpsHealthy)}>
                {isEn ? "Local HTTPS connection" : "Kết nối HTTPS cục bộ"}
              </StatusLine>
              <StatusLine ok={Boolean(status?.companionCertificateMatches)}>
                {isEn ? "Certificate" : "Giấy chứng nhận"}
              </StatusLine>
              <StatusLine ok={Boolean(status?.companionProtocolMatches)}>
                {isEn ? "Protocol version" : "Phiên bản giao thức"}
              </StatusLine>
              <StatusLine ok={Boolean(status?.wordConnected)}>
                {isEn ? "Word connection" : "Kết nối từ"}
              </StatusLine>
              <StatusLine ok={Boolean(status?.powerpointConnected)}>
                {isEn ? "PowerPoint connection" : "Kết nối PowerPoint"}
              </StatusLine>
            </article>
          </div>

          <div className="office-diagnostic-meta">
            <span>
              {isEn ? "Backend" : "Phần cuối"}: {status?.activeBackend ?? "—"}
            </span>
            <span>
              {isEn ? "Companion" : "Người bạn đồng hành"}: {companion?.running
                ? `${companion.bindAddress}:${companion.port}`
                : isEn
                  ? "Stopped"
                  : "Đã dừng"}
            </span>
          </div>

          <div className="office-secondary-actions">
            <button
              type="button"
              className="secondary-button"
              disabled={busy !== null || !hasInstalledComponents}
              onClick={openRuntimeVerification}
            >
              <CheckCircle2 size={15} />
              {isEn ? "Verify Office connection" : "Xác minh kết nối Office"}
            </button>
            <button
              type="button"
              className="secondary-button"
              disabled={busy !== null || !hasInstalledComponents}
              onClick={() =>
                void run("repair", "repair_windows_office_integration")
              }
            >
              <Wrench size={15} />
              {isEn ? "Repair integration" : "Tích hợp sửa chữa"}
            </button>
            <button
              type="button"
              className="secondary-button"
              disabled={busy !== null}
              onClick={() => void run("open-logs", "open_windows_office_logs")}
            >
              <FolderOpen size={15} />
              {isEn ? "Open logs" : "Nhật ký mở"}
            </button>
            <button
              type="button"
              className="secondary-button"
              disabled={busy !== null}
              onClick={() =>
                void run("background-start", "set_office_background_start", {
                  enabled: !status?.backgroundStartEnabled,
                })
              }
            >
              {status?.backgroundStartEnabled ? (
                <ToggleRight size={15} />
              ) : (
                <ToggleLeft size={15} />
              )}
              {status?.backgroundStartEnabled
                ? isEn
                  ? "Disable startup"
                  : "Tắt khởi động"
                : isEn
                  ? "Enable startup"
                  : "Kích hoạt tính năng khởi động"}
            </button>
            <button
              type="button"
              className="secondary-button"
              disabled={busy !== null || Boolean(companion?.running)}
              onClick={() => void run("start", "start_office_companion")}
            >
              <Play size={15} />
              {isEn ? "Start service" : "Bắt đầu dịch vụ"}
            </button>
            <button
              type="button"
              className="secondary-button"
              disabled={busy !== null || !companion?.running}
              onClick={() => void run("stop", "stop_office_companion")}
            >
              <Square size={14} />
              {isEn ? "Stop service" : "Dừng dịch vụ"}
            </button>
            <button
              type="button"
              className="secondary-button danger-subtle"
              disabled={busy !== null || !hasInstalledComponents}
              onClick={() => setConfirmUninstall(true)}
            >
              <Trash2 size={15} />
              {isEn ? "Uninstall integration" : "Gỡ cài đặt tích hợp"}
            </button>
          </div>
        </div>
      </details>

      {confirmRuntimeTest && (
        <div
          className="office-confirm-backdrop"
          role="presentation"
          onMouseDown={() => {
            if (busy !== null) return;
            setConfirmRuntimeTest(false);
            setForceCloseOffice(false);
          }}
        >
          <section
            className="office-confirm-dialog office-runtime-confirm-dialog"
            role="alertdialog"
            aria-modal="true"
            aria-labelledby="office-runtime-test-title"
            aria-describedby="office-runtime-test-description"
            onMouseDown={(event) => event.stopPropagation()}
          >
            <header>
              <span
                className="office-confirm-icon is-verification"
                aria-hidden="true"
              >
                <CheckCircle2 size={19} />
              </span>
              <div>
                <strong id="office-runtime-test-title">
                  {isEn ? "Verify Office connection" : "Xác minh kết nối Office"}
                </strong>
                <p id="office-runtime-test-description">
                  {isEn
                    ? "VisualTeX will start Word and PowerPoint briefly to verify that both add-ins connect. Save your documents and close every Office application before continuing."
                    : "VisualTeX sẽ khởi động nhanh Word và PowerPoint để xác minh rằng cả hai phần bổ trợ đều kết nối. Lưu tài liệu của bạn và đóng mọi ứng dụng Office trước khi tiếp tục."}
                </p>
              </div>
              <button
                type="button"
                className="icon-button compact"
                disabled={busy !== null}
                onClick={() => {
                  setConfirmRuntimeTest(false);
                  setForceCloseOffice(false);
                }}
                aria-label={isEn ? "Cancel verification" : "Hủy xác minh"}
              >
                <X size={16} />
              </button>
            </header>

            <label className="office-force-close-option">
              <input
                type="checkbox"
                checked={forceCloseOffice}
                disabled={busy !== null}
                onChange={(event) => setForceCloseOffice(event.target.checked)}
              />
              <span>
                <strong>
                  {isEn
                    ? "Force-close running Word and PowerPoint"
                    : "Buộc đóng Word và PowerPoint đang chạy"}
                </strong>
                <small>
                  {isEn
                    ? "Use only after saving. Unsaved Office changes may be lost."
                    : "Chỉ sử dụng sau khi lưu. Những thay đổi Office chưa được lưu có thể bị mất."}
                </small>
              </span>
            </label>

            {message && (
              <div className="office-confirm-inline-warning" role="alert">
                <ShieldAlert size={15} />
                <span>{message}</span>
              </div>
            )}

            <footer>
              <button
                type="button"
                className="secondary-button"
                disabled={busy !== null}
                onClick={() => {
                  setConfirmRuntimeTest(false);
                  setForceCloseOffice(false);
                }}
              >
                {isEn ? "Cancel" : "Hủy bỏ"}
              </button>
              <button
                type="button"
                className="primary-button"
                disabled={busy !== null}
                onClick={() => void verifyRuntime()}
              >
                <CheckCircle2 size={15} />
                {busy === "runtime-test"
                  ? isEn
                    ? "Verifying…"
                    : "Đang xác minh…"
                  : forceCloseOffice
                    ? isEn
                      ? "Force close and verify"
                      : "Buộc đóng và xác minh"
                    : isEn
                      ? "I have closed Office; verify"
                      : "Tôi đã đóng Office; xác minh"}
              </button>
            </footer>
          </section>
        </div>
      )}

      {confirmUninstall && (
        <div
          className="office-confirm-backdrop"
          role="presentation"
          onMouseDown={() => busy === null && setConfirmUninstall(false)}
        >
          <section
            className="office-confirm-dialog"
            role="alertdialog"
            aria-modal="true"
            aria-labelledby="office-uninstall-title"
            aria-describedby="office-uninstall-description"
            onMouseDown={(event) => event.stopPropagation()}
          >
            <header>
              <span className="office-confirm-icon" aria-hidden="true">
                <Trash2 size={19} />
              </span>
              <div>
                <strong id="office-uninstall-title">
                  {isEn ? "Uninstall Office integration?" : "Gỡ cài đặt tích hợp Office?"}
                </strong>
                <p id="office-uninstall-description">
                  {isEn
                    ? "This removes the VisualTeX add-ins and OLE registration from Word and PowerPoint. Your formulas and VisualTeX documents will not be deleted."
                    : "Thao tác này sẽ xóa phần bổ trợ VisualTeX và đăng ký OLE khỏi Word và PowerPoint. Công thức và tài liệu VisualTeX của bạn sẽ không bị xóa."}
                </p>
              </div>
              <button
                type="button"
                className="icon-button compact"
                disabled={busy !== null}
                onClick={() => setConfirmUninstall(false)}
                aria-label={isEn ? "Cancel uninstall" : "Hủy gỡ cài đặt"}
              >
                <X size={16} />
              </button>
            </header>
            <footer>
              <button
                type="button"
                className="secondary-button"
                disabled={busy !== null}
                onClick={() => setConfirmUninstall(false)}
              >
                {isEn ? "Cancel" : "Hủy bỏ"}
              </button>
              <button
                type="button"
                className="danger-button"
                disabled={busy !== null}
                onClick={() => void uninstall()}
              >
                <Trash2 size={15} />
                {busy === "uninstall-ole"
                  ? isEn
                    ? "Uninstalling…"
                    : "Đang gỡ cài đặt…"
                  : isEn
                    ? "Uninstall"
                    : "Gỡ cài đặt"}
              </button>
            </footer>
          </section>
        </div>
      )}
    </section>
  );
}
