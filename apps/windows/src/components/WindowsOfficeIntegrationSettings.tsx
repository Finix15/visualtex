import { useCallback, useEffect, useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import {
  CheckCircle2,
  Download,
  ExternalLink,
  FolderOpen,
  Play,
  RefreshCw,
  ShieldAlert,
  Square,
  ToggleLeft,
  ToggleRight,
  Trash2,
  Wrench,
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
      } catch (error) {
        setMessage(
          errorMessage(
            error,
            isEn ? "Windows Office operation failed." : "Windows Office 操作失败。",
          ),
        );
      } finally {
        setBusy(null);
      }
    },
    [isEn, refresh],
  );

  return (
    <section className="settings-section office-integration-section">
      <div className="settings-section-heading office-settings-heading">
        <div>
          <strong>{isEn ? "Windows Office integration" : "Windows Office 集成"}</strong>
          <p>
            {isEn
              ? "Windows Office integration uses native Word/PowerPoint Ribbon COM add-ins together with the VisualTeX Formula OLE LocalServer. Legacy Office.js Trusted Catalog manifests are not used."
              : "Windows Office 集成统一使用 Word/PowerPoint 原生 Ribbon COM 加载项与 VisualTeX Formula OLE LocalServer，不再使用旧 Office.js Trusted Catalog 清单。"}
          </p>
        </div>
        <button
          type="button"
          className="icon-button compact"
          onClick={() => void refresh()}
          disabled={busy !== null}
          title={isEn ? "Refresh" : "刷新"}
        >
          <RefreshCw size={15} className={busy === "refresh" ? "is-spinning" : ""} />
        </button>
      </div>

      <div className="office-status-grid">
        <article className="office-status-card">
          <header>
            <strong>{isEn ? "Current backend" : "当前后端"}</strong>
            <StatusLine ok={Boolean(status && !status.lastError)}>
              {status?.activeBackend ?? "—"}
            </StatusLine>
          </header>
          <dl>
            <div><dt>{isEn ? "Selected mode" : "设置模式"}</dt><dd>{isEn ? "Native Ribbon + OLE LocalServer" : "原生 Ribbon + OLE LocalServer"}</dd></div>
            <div><dt>{isEn ? "Files and registry" : "文件与注册"}</dt><dd>{status?.staticInstallVerified ? (isEn ? "Verified" : "已验证") : (isEn ? "Incomplete" : "不完整")}</dd></div>
            <div><dt>{isEn ? "Companion runtime" : "伴侣服务运行时"}</dt><dd>{status?.officeRuntimeVerified ? (isEn ? "Verified" : "已验证") : (isEn ? "Not verified" : "尚未验证")}</dd></div>
          </dl>
        </article>

        <article className="office-status-card">
          <header><strong>{isEn ? "Native Office components" : "原生 Office 组件"}</strong></header>
          <StatusLine ok={Boolean(status?.wordFilesPresent)}>
            {isEn ? "Word add-in files present" : "Word 加载项文件存在"}
          </StatusLine>
          <StatusLine ok={Boolean(status?.wordRegistryComplete)}>
            {isEn ? "Word COM registration complete" : "Word COM 注册完整"}
          </StatusLine>
          <StatusLine ok={Boolean(status?.wordLoadEnabled)}>
            Word LoadBehavior=3
          </StatusLine>
          <StatusLine ok={Boolean(status?.powerpointFilesPresent)}>
            {isEn ? "PowerPoint add-in files present" : "PowerPoint 加载项文件存在"}
          </StatusLine>
          <StatusLine ok={Boolean(status?.powerpointRegistryComplete)}>
            {isEn ? "PowerPoint COM registration complete" : "PowerPoint COM 注册完整"}
          </StatusLine>
          <StatusLine ok={Boolean(status?.powerpointLoadEnabled)}>
            PowerPoint LoadBehavior=3
          </StatusLine>
          <StatusLine ok={Boolean(status?.oleLocalServerHealthy)}>
            {isEn ? "Formula OLE LocalServer healthy" : "公式 OLE LocalServer 健康"}
          </StatusLine>
        </article>

        <article className="office-status-card">
          <header><strong>{isEn ? "Runtime and security" : "运行时与安全"}</strong></header>
          <StatusLine ok={Boolean(status?.companionProcessRunning)}>
            {isEn ? "Companion process running" : "伴侣进程正在运行"}
          </StatusLine>
          <StatusLine ok={Boolean(status?.companionPortListening)}>
            {isEn ? "Companion TCP port listening" : "伴侣 TCP 端口正在监听"}
          </StatusLine>
          <StatusLine ok={Boolean(status?.companionHttpsHealthy)}>
            {isEn ? "Companion HTTPS health passed" : "伴侣 HTTPS 健康检查通过"}
          </StatusLine>
          <StatusLine ok={Boolean(status?.companionCertificateMatches)}>
            {isEn ? "Companion certificate matches" : "伴侣证书一致"}
          </StatusLine>
          <StatusLine ok={Boolean(status?.companionProtocolMatches)}>
            {isEn ? "Companion protocol matches" : "伴侣协议版本一致"}
          </StatusLine>
          <StatusLine ok={Boolean(status?.wordConnected)}>
            Word COMAddIn.Connect={String(Boolean(status?.wordConnected))}
          </StatusLine>
          <StatusLine ok={Boolean(status?.powerpointConnected)}>
            PowerPoint COMAddIn.Connect={String(Boolean(status?.powerpointConnected))}
          </StatusLine>
          <StatusLine ok={Boolean(status?.currentUserCertificateTrusted)}>
            {isEn ? "Current-user HTTPS certificate trusted" : "当前用户 HTTPS 证书受信任"}
          </StatusLine>
        </article>

        <article className="office-status-card office-status-card-wide">
          <header><strong>{isEn ? "Session companion" : "Session 伴侣服务"}</strong></header>
          <StatusLine ok={Boolean(companion?.running)}>
            {companion?.running
              ? `${companion.bindAddress}:${companion.port}`
              : isEn ? "Stopped" : "已停止"}
          </StatusLine>
          <p>
            {isEn
              ? "OCR, formula rendering, cache and editor sessions remain in the shared HTTPS companion."
              : "OCR、公式渲染、缓存和编辑会话继续由共享 HTTPS 伴侣服务负责。"}
          </p>
        </article>
      </div>

      {(message || status?.lastError || companion?.lastError) && (
        <div className="office-settings-warning" role="alert">
          <ShieldAlert size={15} />
          <pre className="office-settings-diagnostic">
            {message || status?.lastError || companion?.lastError}
          </pre>
        </div>
      )}

      <div className="office-settings-actions">
        <button
          type="button"
          className="primary-button"
          disabled={busy !== null}
          onClick={() => void run("install-ole", "install_windows_ole_integration")}
        >
          <Download size={15} />
          {isEn ? "Install native Office integration" : "安装原生 Office 集成"}
        </button>
        <button
          type="button"
          className="secondary-button"
          disabled={busy !== null}
          onClick={() => void run("repair", "repair_windows_office_integration")}
        >
          <Wrench size={15} />
          {isEn ? "Repair native Office integration" : "修复原生 Office 集成"}
        </button>
        <button
          type="button"
          className="secondary-button"
          disabled={busy !== null}
          onClick={() => void run("runtime-test", "test_windows_office_runtime")}
        >
          <CheckCircle2 size={15} />
          {isEn ? "Verify Office runtime" : "验证 Office 运行时"}
        </button>
        <button
          type="button"
          className="secondary-button"
          disabled={busy !== null}
          onClick={() => void run("open-logs", "open_windows_office_logs")}
        >
          <FolderOpen size={15} />
          {isEn ? "Open diagnostic logs" : "打开诊断日志"}
        </button>
        <button
          type="button"
          className="secondary-button"
          disabled={busy !== null}
          onClick={() => void run(
            "background-start",
            "set_office_background_start",
            { enabled: !status?.backgroundStartEnabled },
          )}
        >
          {status?.backgroundStartEnabled ? <ToggleRight size={15} /> : <ToggleLeft size={15} />}
          {status?.backgroundStartEnabled
            ? isEn ? "Disable startup" : "关闭开机启动"
            : isEn ? "Enable startup" : "启用开机启动"}
        </button>
        <button
          type="button"
          className="secondary-button danger-subtle"
          disabled={busy !== null}
          onClick={() => void run("uninstall-ole", "uninstall_windows_ole_integration")}
        >
          <Trash2 size={15} />
          {isEn ? "Remove Office integration" : "移除 Office 集成"}
        </button>
        <button
          type="button"
          className="secondary-button"
          disabled={busy !== null || Boolean(companion?.running)}
          onClick={() => void run("start", "start_office_companion")}
        >
          <Play size={15} />
          {isEn ? "Start companion" : "启动伴侣服务"}
        </button>
        <button
          type="button"
          className="secondary-button"
          disabled={busy !== null || !companion?.running}
          onClick={() => void run("stop", "stop_office_companion")}
        >
          <Square size={14} />
          {isEn ? "Stop companion" : "停止伴侣服务"}
        </button>
        <button type="button" className="secondary-button" disabled={busy !== null} onClick={() => void run("word", "open_word")}>
          <ExternalLink size={15} />{isEn ? "Open Word" : "打开 Word"}
        </button>
        <button type="button" className="secondary-button" disabled={busy !== null} onClick={() => void run("powerpoint", "open_powerpoint")}>
          <ExternalLink size={15} />{isEn ? "Open PowerPoint" : "打开 PowerPoint"}
        </button>
      </div>
    </section>
  );
}
