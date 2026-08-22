import { useEffect, useRef, useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import {
  CheckCircle2,
  Download,
  ExternalLink,
  FileText,
  LoaderCircle,
  Presentation,
  RefreshCw,
  ShieldAlert,
} from "lucide-react";
import type { Language } from "../stores/editorStore";
import { PowerPointAddinGuide } from "./PowerPointAddinGuide";

interface MacOfflineHostStatus {
  applicationInstalled: boolean;
  applicationRunning: boolean;
  filesPresent: boolean;
  filesInstalled: boolean;
  loaded: boolean;
  pluginVersion: string | null;
  installPaths: string[];
  healthPath: string;
  lastError: string | null;
}

interface MacOfflineOfficeStatus {
  word: MacOfflineHostStatus;
  powerpoint: MacOfflineHostStatus;
  compiledArtifactsAvailable: boolean;
  resourceRoot: string;
  powerpointAddinPath: string;
  wordScriptPath: string;
  powerpointScriptPath: string;
  tutorialPath: string;
}

interface Props {
  open: boolean;
  language: Language;
  mode?: "setup" | "update" | "repair";
  powerpointRegistrationRequired?: boolean;
  onComplete: (installed: boolean) => void;
}

function messageFrom(error: unknown, fallback: string) {
  if (error instanceof Error && error.message) return error.message;
  if (typeof error === "string" && error.trim()) return error;
  return fallback;
}

function wait(milliseconds: number) {
  return new Promise<void>((resolve) => window.setTimeout(resolve, milliseconds));
}

export function MacOfficeFirstRunPrompt({
  open,
  language,
  mode = "setup",
  powerpointRegistrationRequired = false,
  onComplete,
}: Props) {
  const [status, setStatus] = useState<MacOfflineOfficeStatus | null>(null);
  const [busy, setBusy] = useState<string | null>(null);
  const [error, setError] = useState("");
  const dialogRef = useRef<HTMLElement>(null);
  const isEn = language === "en";

  const refresh = async () => {
    const next = await invoke<MacOfflineOfficeStatus>(
      "get_macos_offline_office_install_status",
    );
    setStatus(next);
    return next;
  };

  useEffect(() => {
    if (!open) return;
    setError("");
    setBusy("refresh");
    void refresh()
      .catch((reason) => {
        setError(
          messageFrom(
            reason,
            isEn
              ? "Unable to inspect the native Office add-ins on this Mac."
              : "Không thể kiểm tra các phần bổ trợ Office gốc trên máy Mac này.",
          ),
        );
      })
      .finally(() => setBusy(null));
    const frame = window.requestAnimationFrame(() => {
      dialogRef.current?.querySelector<HTMLElement>("button")?.focus();
    });
    return () => window.cancelAnimationFrame(frame);
  }, [isEn, open]);

  if (!open) return null;

  const officeDetected = Boolean(
    status?.word.applicationInstalled || status?.powerpoint.applicationInstalled,
  );
  const nativeFilesReady = Boolean(
    status?.compiledArtifactsAvailable &&
      (!status.word.applicationInstalled || status.word.filesInstalled) &&
      (!status.powerpoint.applicationInstalled || status.powerpoint.filesInstalled),
  );
  const powerpointNeedsRegistration = Boolean(
    powerpointRegistrationRequired &&
      status?.powerpoint.applicationInstalled &&
      status.powerpoint.filesInstalled &&
      !status.powerpoint.loaded,
  );
  const officeHostsRunning = Boolean(
    status?.word.applicationRunning || status?.powerpoint.applicationRunning,
  );
  const updateRequired = Boolean(
    mode !== "setup" &&
      ((status?.word.applicationInstalled && !status.word.filesInstalled) ||
        (status?.powerpoint.applicationInstalled &&
          !status.powerpoint.filesInstalled)),
  );

  const install = async () => {
    setBusy("install");
    setError("");
    try {
      const next = await invoke<MacOfflineOfficeStatus>(
        "install_macos_offline_office_addins",
      );
      setStatus(next);
      if (
        mode !== "setup" &&
        (!next.word.applicationInstalled || next.word.filesInstalled) &&
        (!next.powerpoint.applicationInstalled || next.powerpoint.filesInstalled)
      ) {
        onComplete(true);
      }
    } catch (reason) {
      setError(
        messageFrom(
          reason,
          isEn
            ? "VisualTeX could not install the native Word and PowerPoint add-ins."
            : "VisualTeX không thể cài đặt các phần bổ trợ Word và PowerPoint gốc.",
        ),
      );
    } finally {
      setBusy(null);
    }
  };

  const quitOfficeAndUpdate = async () => {
    setBusy("quit-update");
    setError("");
    try {
      await invoke("request_quit_macos_office_hosts_for_addin_update");
      const deadline = Date.now() + 120_000;
      while (Date.now() < deadline) {
        await wait(500);
        const current = await refresh();
        if (!current.word.applicationRunning && !current.powerpoint.applicationRunning) {
          const next = await invoke<MacOfflineOfficeStatus>(
            "install_macos_offline_office_addins",
          );
          setStatus(next);
          if (
            (!next.word.applicationInstalled || next.word.filesInstalled) &&
            (!next.powerpoint.applicationInstalled || next.powerpoint.filesInstalled)
          ) {
            onComplete(true);
          }
          return;
        }
      }
      throw new Error(
        isEn
          ? "Timed out waiting for Word and PowerPoint to quit. Finish any Save prompts, then try again."
          : "Đã hết thời gian chờ Word và PowerPoint thoát. Hãy hoàn tất mọi lời nhắc Lưu rồi thử lại.",
      );
    } catch (reason) {
      setError(
        messageFrom(
          reason,
          isEn
            ? "VisualTeX could not finish the Office add-in update."
            : "VisualTeX không thể hoàn tất quá trình cập nhật phần bổ trợ Office.",
        ),
      );
    } finally {
      setBusy(null);
    }
  };

  const runAction = async (name: string, command: string) => {
    setBusy(name);
    setError("");
    try {
      await invoke(command);
      if (command === "open_powerpoint") {
        window.setTimeout(() => void refresh().catch(() => undefined), 1200);
      }
    } catch (reason) {
      setError(
        messageFrom(
          reason,
          isEn ? "The requested Office action failed." : "Hành động Office được yêu cầu không thành công.",
        ),
      );
    } finally {
      setBusy(null);
    }
  };

  return (
    <div className="office-first-run-backdrop">
      <section
        ref={dialogRef}
        className="office-first-run-dialog is-native-office"
        role="dialog"
        aria-modal="true"
        aria-labelledby="office-first-run-title"
      >
        <header>
          <span><Download size={20} /></span>
          <div>
            <strong id="office-first-run-title">
              {mode === "update"
                ? isEn
                  ? "Update the VisualTeX Office add-ins"
                  : "Cập nhật phần bổ trợ VisualTeX Office"
                : mode === "repair"
                  ? isEn
                    ? "Repair the VisualTeX Office add-ins"
                    : "Sửa chữa phần bổ trợ VisualTeX Office"
                  : isEn
                    ? "Set up VisualTeX for Word and PowerPoint"
                    : "Thiết lập VisualTeX cho Word và PowerPoint"}
            </strong>
            <p>
              {mode === "update"
                ? isEn
                  ? "The installed DOTM or PPAM belongs to an older VisualTeX build. You do not need to delete it manually; VisualTeX will replace the old add-ins with the versions bundled in this app."
                  : "DOTM hoặc PPAM được cài đặt thuộc về bản dựng VisualTeX cũ hơn. Bạn không cần phải xóa nó theo cách thủ công; VisualTeX sẽ thay thế các phần bổ trợ cũ bằng các phiên bản đi kèm trong ứng dụng này."
                : mode === "repair"
                  ? isEn
                    ? "VisualTeX detected a missing or incomplete Office add-in installation. It will restore the required DOTM/PPAM files without asking you to re-register an already configured PowerPoint add-in."
                    : "VisualTeX đã phát hiện thấy bản cài đặt phần bổ trợ Office bị thiếu hoặc chưa hoàn chỉnh. Nó sẽ khôi phục các tệp DOTM/PPAM cần thiết mà không yêu cầu bạn đăng ký lại phần bổ trợ PowerPoint đã được định cấu hình."
                  : isEn
                    ? "VisualTeX installs a Word DOTM template and a PowerPoint PPAM add-in. Both run locally and open the desktop formula editor when needed."
                    : "VisualTeX cài đặt mẫu Word DOTM và phần bổ trợ PowerPoint PPAM. Cả hai đều chạy cục bộ và mở trình chỉnh sửa công thức trên máy tính để bàn khi cần."}
            </p>
          </div>
        </header>

        <div className="office-first-run-hosts">
          <article className={status?.word.filesInstalled && status.word.loaded ? "is-loaded" : status?.word.filesInstalled ? "is-files-ready" : ""}>
            <FileText size={20} />
            <div>
              <strong>Microsoft Word · DOTM</strong>
              <small>
                {!status?.word.applicationInstalled
                  ? isEn ? "Word not detected" : "Không phát hiện được từ"
                  : status.word.filesInstalled && status.word.loaded
                    ? isEn ? "Installed and loaded" : "Đã cài đặt và tải"
                    : status.word.filesInstalled
                      ? isEn ? "Installed; restart Word" : "Đã cài đặt; khởi động lại Word"
                      : status.word.filesPresent
                        ? isEn ? "Older add-in detected; update required" : "Đã phát hiện thấy phần bổ trợ cũ hơn; yêu cầu cập nhật"
                        : isEn ? "Not installed" : "Chưa cài đặt"}
              </small>
            </div>
            {status?.word.filesInstalled && status.word.loaded ? <CheckCircle2 size={17} /> : status?.word.filesInstalled ? <ShieldAlert size={17} /> : null}
          </article>
          <article className={status?.powerpoint.filesInstalled && status.powerpoint.loaded ? "is-loaded" : status?.powerpoint.filesInstalled ? "is-files-ready" : ""}>
            <Presentation size={20} />
            <div>
              <strong>Microsoft PowerPoint · PPAM</strong>
              <small>
                {!status?.powerpoint.applicationInstalled
                  ? isEn ? "PowerPoint not detected" : "Không tìm thấy PowerPoint"
                  : status.powerpoint.filesInstalled && status.powerpoint.loaded
                    ? isEn ? "Installed and loaded" : "Đã cài đặt và tải"
                    : status.powerpoint.filesInstalled
                      ? powerpointRegistrationRequired
                        ? isEn ? "Installed; register once" : "Đã cài đặt; đăng ký một lần"
                        : isEn ? "Installed; registration is preserved" : "Đã cài đặt; đăng ký được bảo tồn"
                      : status.powerpoint.filesPresent
                        ? isEn ? "Older add-in detected; update required" : "Đã phát hiện thấy phần bổ trợ cũ hơn; yêu cầu cập nhật"
                        : isEn ? "Not installed" : "Chưa cài đặt"}
              </small>
            </div>
            {status?.powerpoint.filesInstalled && status.powerpoint.loaded ? <CheckCircle2 size={17} /> : status?.powerpoint.filesInstalled ? <ShieldAlert size={17} /> : null}
          </article>
        </div>

        {powerpointNeedsRegistration ? (
          <div className="office-first-run-powerpoint-guide">
            <div className="office-first-run-note is-important">
              <p>
                {isEn
                  ? "The PPAM file is ready, but PowerPoint has not registered it. Not seeing VisualTeX in the Add-ins list yet is expected. Click + first, then choose the PPAM file."
                  : "File PPAM đã sẵn sàng nhưng PowerPoint chưa đăng ký. Dự kiến ​​sẽ không thấy VisualTeX trong danh sách Bổ trợ. Trước tiên hãy nhấp vào +, sau đó chọn tệp PPAM."}
              </p>
            </div>
            <PowerPointAddinGuide language={language} compact loaded={false} />
            <div className="office-first-run-guide-actions">
              <button
                type="button"
                className="secondary-button"
                disabled={busy !== null}
                onClick={() => void runAction("reveal", "reveal_macos_powerpoint_addin")}
              >
                <ExternalLink size={15} />
                {isEn ? "Show PPAM in Finder" : "Hiển thị PPAM trong Finder"}
              </button>
              <button
                type="button"
                className="secondary-button"
                disabled={busy !== null}
                onClick={() => void runAction("powerpoint", "open_powerpoint")}
              >
                <Presentation size={15} />
                {isEn ? "Open PowerPoint" : "Mở PowerPoint"}
              </button>
              <button
                type="button"
                className="secondary-button"
                disabled={busy !== null}
                onClick={() => {
                  setBusy("refresh");
                  void refresh().finally(() => setBusy(null));
                }}
              >
                <RefreshCw size={15} className={busy === "refresh" ? "is-spinning" : ""} />
                {isEn ? "Refresh status" : "Trạng thái làm mới"}
              </button>
            </div>
          </div>
        ) : (
          <div className="office-first-run-note">
            <p>
              {mode === "update"
                ? isEn
                  ? "VisualTeX updates the existing DOTM and PPAM in place. PowerPoint keeps the same registered PPAM path, so an update does not require registering the add-in again."
                  : "VisualTeX cập nhật DOTM và PPAM hiện có. PowerPoint giữ nguyên đường dẫn PPAM đã đăng ký, do đó, bản cập nhật không yêu cầu đăng ký lại phần bổ trợ."
                : mode === "repair"
                  ? isEn
                    ? "VisualTeX will restore the missing Office files. Any PowerPoint PPAM that was already registered keeps the same path and does not need to be registered again."
                    : "VisualTeX sẽ khôi phục các tệp Office bị thiếu. Bất kỳ PPAM PowerPoint nào đã được đăng ký đều giữ nguyên đường dẫn và không cần phải đăng ký lại."
                  : powerpointRegistrationRequired
                    ? isEn
                      ? "Word loads VisualTeX automatically from its Startup folder after Word restarts. PowerPoint needs one manual registration through Tools → PowerPoint Add-ins; later updates keep the same PPAM path."
                      : "Word tự động tải VisualTeX từ thư mục Khởi động sau khi Word khởi động lại. PowerPoint cần một lần đăng ký thủ công thông qua Công cụ → Phần bổ trợ PowerPoint; các bản cập nhật sau này giữ nguyên đường dẫn PPAM."
                    : isEn
                      ? "VisualTeX will install the required Office files. Existing PowerPoint registration is preserved."
                      : "VisualTeX sẽ cài đặt các tệp Office cần thiết. Đăng ký PowerPoint hiện tại được giữ nguyên."}
            </p>
            {mode !== "setup" && updateRequired && officeHostsRunning && (
              <p className="is-warning">
                {isEn
                  ? "Save your Office documents first. VisualTeX must let Word and PowerPoint quit normally before replacing loaded VBA add-ins. Click the update button below; any unsaved Office document will still receive its normal Save prompt."
                  : "Lưu tài liệu Office của bạn trước. VisualTeX phải để Word và PowerPoint thoát bình thường trước khi thay thế các phần bổ trợ VBA đã tải. Nhấp vào nút cập nhật bên dưới; mọi tài liệu Office chưa được lưu sẽ vẫn nhận được lời nhắc Lưu thông thường."}
              </p>
            )}
            {!officeDetected && status && (
              <p className="is-warning">
                {isEn
                  ? "Word or PowerPoint was not found. Open the Office application once, then return to Settings to install the native add-ins."
                  : "Không tìm thấy Word hoặc PowerPoint. Mở ứng dụng Office một lần, sau đó quay lại Cài đặt để cài đặt các phần bổ trợ gốc."}
              </p>
            )}
            {nativeFilesReady && (
              <p className="is-warning">
                {powerpointRegistrationRequired
                  ? isEn
                    ? "The files are installed. Restart Word, then register the PowerPoint PPAM once to finish first-time setup."
                    : "Các tập tin đã được cài đặt. Khởi động lại Word, sau đó đăng ký PPAM PowerPoint một lần để hoàn tất quá trình thiết lập lần đầu."
                  : isEn
                    ? "The current add-in files are ready. Reopen Word or PowerPoint so Office loads the repaired/updated files."
                    : "Các tập tin bổ trợ hiện tại đã sẵn sàng. Mở lại Word hoặc PowerPoint để Office tải các tệp đã sửa chữa/cập nhật."}
              </p>
            )}
            {error && <p className="is-warning" role="alert">{error}</p>}
          </div>
        )}

        {powerpointNeedsRegistration && error && (
          <div className="office-settings-warning" role="alert">
            <ShieldAlert size={15} />
            <span>{error}</span>
          </div>
        )}

        <footer>
          <button
            type="button"
            className="secondary-button"
            disabled={busy !== null}
            onClick={() => onComplete(false)}
          >
            {isEn ? "Later" : "Sau đó"}
          </button>
          <button
            type="button"
            className="primary-button"
            disabled={busy !== null || (!nativeFilesReady && !officeDetected)}
            onClick={() => {
              if (nativeFilesReady) {
                onComplete(true);
              } else if (mode !== "setup" && officeHostsRunning) {
                void quitOfficeAndUpdate();
              } else {
                void install();
              }
            }}
          >
            {busy === "install" || busy === "quit-update" ? (
              <LoaderCircle className="is-spinning" size={16} />
            ) : nativeFilesReady ? (
              <CheckCircle2 size={16} />
            ) : (
              <Download size={16} />
            )}
            {nativeFilesReady
              ? mode === "setup"
                ? isEn ? "Continue" : "Tiếp tục"
                : isEn ? "Done" : "Xong"
              : mode === "update"
                ? officeHostsRunning
                  ? isEn ? "Quit Office and update add-ins" : "Thoát khỏi Office và cập nhật phần bổ trợ"
                  : isEn ? "Update DOTM and PPAM" : "Cập nhật DOTM và PPAM"
                : mode === "repair"
                  ? officeHostsRunning
                    ? isEn ? "Quit Office and repair add-ins" : "Thoát khỏi Office và sửa chữa các phần bổ trợ"
                    : isEn ? "Repair DOTM and PPAM" : "Sửa chữa DOTM và PPAM"
                  : isEn ? "Install DOTM and PPAM" : "Cài đặt DOTM và PPAM"}
          </button>
        </footer>
      </section>
    </div>
  );
}
