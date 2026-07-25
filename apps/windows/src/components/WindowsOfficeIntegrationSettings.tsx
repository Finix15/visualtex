import { useCallback, useEffect, useMemo, useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import {
  CheckCircle2,
  ChevronDown,
  CircleAlert,
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
            : "无法读取 Windows Office 集成状态。",
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
            isEn ? "Windows Office operation failed." : "Windows Office 操作失败。",
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
  const integrationReady = installHealthy && runtimeHealthy && !status?.lastError;
  const hasInstalledComponents = Boolean(
    status?.wordFilesPresent ||
      status?.powerpointFilesPresent ||
      status?.wordRegistryComplete ||
      status?.powerpointRegistryComplete ||
      status?.oleLocalServerHealthy,
  );
  const statusCopy = useMemo(() => {
    if (!status) {
      return {
        title: isEn ? "Checking Office integration…" : "正在检查 Office 集成…",
        description: isEn
          ? "VisualTeX is reading the Word and PowerPoint add-in status."
          : "VisualTeX 正在读取 Word 和 PowerPoint 加载项状态。",
      };
    }
    if (integrationReady) {
      return {
        title: isEn ? "Office integration is ready" : "Office 集成可正常使用",
        description: isEn
          ? "Word and PowerPoint can create and edit VisualTeX formulas."
          : "Word 和 PowerPoint 已可创建、插入和编辑 VisualTeX 公式。",
      };
    }
    if (hasInstalledComponents) {
      return {
        title: isEn ? "Office integration needs repair" : "Office 集成需要修复",
        description: isEn
          ? "Some components are installed, but the complete runtime check did not pass."
          : "部分组件已经安装，但完整运行检查尚未通过。",
      };
    }
    return {
      title: isEn ? "Office integration is not installed" : "尚未安装 Office 集成",
      description: isEn
        ? "Install once to add VisualTeX tools to Word and PowerPoint."
        : "安装后即可在 Word 和 PowerPoint 中直接使用 VisualTeX。",
    };
  }, [hasInstalledComponents, integrationReady, isEn, status]);

  const diagnosticMessage = message || status?.lastError || companion?.lastError;

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
          <h3>{isEn ? "Office integration" : "Office 集成"}</h3>
          <p>
            {isEn
              ? "Use VisualTeX directly in Microsoft Word and PowerPoint."
              : "在 Microsoft Word 和 PowerPoint 中直接插入和编辑公式。"}
          </p>
        </div>
        <button
          type="button"
          className="icon-button compact"
          onClick={() => void refresh()}
          disabled={busy !== null}
          aria-label={isEn ? "Refresh Office status" : "刷新 Office 状态"}
          title={isEn ? "Refresh status" : "刷新状态"}
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
            : hasInstalledComponents
              ? "needs-attention"
              : "is-not-installed"
        }`}
      >
        <span className="office-summary-icon" aria-hidden="true">
          {integrationReady ? <CheckCircle2 size={22} /> : <CircleAlert size={22} />}
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
                : "正在安装…"
              : isEn
                ? "Install Office integration"
                : "安装 Office 集成"}
          </button>
        )}
        {(integrationReady || hasInstalledComponents) && (
          <button
            type="button"
            className={integrationReady ? "secondary-button" : "primary-button"}
            disabled={busy !== null}
            onClick={() =>
              void run("repair", "repair_windows_office_integration")
            }
          >
            <Wrench size={16} />
            {busy === "repair"
              ? isEn
                ? "Repairing…"
                : "正在修复…"
              : isEn
                ? "Repair integration"
                : "修复 Office 集成"}
          </button>
        )}
        <button
          type="button"
          className="secondary-button"
          disabled={busy !== null}
          onClick={() => void run("word", "open_word")}
        >
          <ExternalLink size={15} />
          {isEn ? "Open Word" : "打开 Word"}
        </button>
        <button
          type="button"
          className="secondary-button"
          disabled={busy !== null}
          onClick={() => void run("powerpoint", "open_powerpoint")}
        >
          <ExternalLink size={15} />
          {isEn ? "Open PowerPoint" : "打开 PowerPoint"}
        </button>
      </div>

      <details className="office-advanced-settings">
        <summary>
          <span>
            <Settings2 size={15} />
            {isEn ? "Advanced diagnostics" : "高级诊断与维护"}
          </span>
          <ChevronDown size={15} className="office-details-chevron" />
        </summary>

        <div className="office-advanced-content">
          <div className="office-status-grid office-status-grid-compact">
            <article className="office-status-card">
              <header>
                <strong>{isEn ? "Installation" : "安装状态"}</strong>
                <StatusLine ok={installHealthy}>
                  {installHealthy
                    ? isEn
                      ? "Complete"
                      : "完整"
                    : isEn
                      ? "Incomplete"
                      : "不完整"}
                </StatusLine>
              </header>
              <StatusLine ok={Boolean(status?.wordFilesPresent)}>
                {isEn ? "Word add-in files" : "Word 加载项文件"}
              </StatusLine>
              <StatusLine ok={Boolean(status?.wordRegistryComplete)}>
                {isEn ? "Word registration" : "Word 注册信息"}
              </StatusLine>
              <StatusLine ok={Boolean(status?.wordLoadEnabled)}>
                {isEn ? "Word add-in enabled" : "Word 加载项已启用"}
              </StatusLine>
              <StatusLine ok={Boolean(status?.powerpointFilesPresent)}>
                {isEn ? "PowerPoint add-in files" : "PowerPoint 加载项文件"}
              </StatusLine>
              <StatusLine ok={Boolean(status?.powerpointRegistryComplete)}>
                {isEn ? "PowerPoint registration" : "PowerPoint 注册信息"}
              </StatusLine>
              <StatusLine ok={Boolean(status?.powerpointLoadEnabled)}>
                {isEn ? "PowerPoint add-in enabled" : "PowerPoint 加载项已启用"}
              </StatusLine>
              <StatusLine ok={Boolean(status?.oleLocalServerHealthy)}>
                {isEn ? "Formula OLE service" : "公式 OLE 服务"}
              </StatusLine>
            </article>

            <article className="office-status-card">
              <header>
                <strong>{isEn ? "Runtime" : "运行状态"}</strong>
                <StatusLine ok={runtimeHealthy}>
                  {runtimeHealthy
                    ? isEn
                      ? "Available"
                      : "可用"
                    : isEn
                      ? "Unavailable"
                      : "不可用"}
                </StatusLine>
              </header>
              <StatusLine ok={Boolean(status?.companionProcessRunning)}>
                {isEn ? "Companion process" : "伴侣进程"}
              </StatusLine>
              <StatusLine ok={Boolean(status?.companionPortListening)}>
                {isEn ? "Local port" : "本地端口"}
              </StatusLine>
              <StatusLine ok={Boolean(status?.companionHttpsHealthy)}>
                {isEn ? "Local HTTPS connection" : "本地 HTTPS 连接"}
              </StatusLine>
              <StatusLine ok={Boolean(status?.companionCertificateMatches)}>
                {isEn ? "Certificate" : "证书"}
              </StatusLine>
              <StatusLine ok={Boolean(status?.companionProtocolMatches)}>
                {isEn ? "Protocol version" : "协议版本"}
              </StatusLine>
              <StatusLine ok={Boolean(status?.wordConnected)}>
                {isEn ? "Word connection" : "Word 连接"}
              </StatusLine>
              <StatusLine ok={Boolean(status?.powerpointConnected)}>
                {isEn ? "PowerPoint connection" : "PowerPoint 连接"}
              </StatusLine>
            </article>
          </div>

          <div className="office-diagnostic-meta">
            <span>
              {isEn ? "Backend" : "后端"}: {status?.activeBackend ?? "—"}
            </span>
            <span>
              {isEn ? "Companion" : "伴侣服务"}: {companion?.running
                ? `${companion.bindAddress}:${companion.port}`
                : isEn
                  ? "Stopped"
                  : "已停止"}
            </span>
          </div>

          <div className="office-secondary-actions">
            <button
              type="button"
              className="secondary-button"
              disabled={busy !== null}
              onClick={() =>
                void run("runtime-test", "test_windows_office_runtime")
              }
            >
              <CheckCircle2 size={15} />
              {isEn ? "Verify runtime" : "验证运行环境"}
            </button>
            <button
              type="button"
              className="secondary-button"
              disabled={busy !== null}
              onClick={() => void run("open-logs", "open_windows_office_logs")}
            >
              <FolderOpen size={15} />
              {isEn ? "Open logs" : "打开诊断日志"}
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
                  : "关闭开机启动"
                : isEn
                  ? "Enable startup"
                  : "启用开机启动"}
            </button>
            <button
              type="button"
              className="secondary-button"
              disabled={busy !== null || Boolean(companion?.running)}
              onClick={() => void run("start", "start_office_companion")}
            >
              <Play size={15} />
              {isEn ? "Start service" : "启动服务"}
            </button>
            <button
              type="button"
              className="secondary-button"
              disabled={busy !== null || !companion?.running}
              onClick={() => void run("stop", "stop_office_companion")}
            >
              <Square size={14} />
              {isEn ? "Stop service" : "停止服务"}
            </button>
            <button
              type="button"
              className="secondary-button danger-subtle"
              disabled={busy !== null || !hasInstalledComponents}
              onClick={() => setConfirmUninstall(true)}
            >
              <Trash2 size={15} />
              {isEn ? "Uninstall integration" : "卸载 Office 集成"}
            </button>
          </div>
        </div>
      </details>

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
                  {isEn ? "Uninstall Office integration?" : "确定卸载 Office 集成？"}
                </strong>
                <p id="office-uninstall-description">
                  {isEn
                    ? "This removes the VisualTeX add-ins and OLE registration from Word and PowerPoint. Your formulas and VisualTeX documents will not be deleted."
                    : "这会从 Word 和 PowerPoint 中移除 VisualTeX 加载项及 OLE 注册，但不会删除已有公式或 VisualTeX 文档。"}
                </p>
              </div>
              <button
                type="button"
                className="icon-button compact"
                disabled={busy !== null}
                onClick={() => setConfirmUninstall(false)}
                aria-label={isEn ? "Cancel uninstall" : "取消卸载"}
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
                {isEn ? "Cancel" : "取消"}
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
                    : "正在卸载…"
                  : isEn
                    ? "Uninstall"
                    : "确认卸载"}
              </button>
            </footer>
          </section>
        </div>
      )}
    </section>
  );
}
