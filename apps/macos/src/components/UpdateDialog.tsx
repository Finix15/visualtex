import { useEffect, useMemo, useRef } from "react";
import {
  CheckCircle2,
  Download,
  LoaderCircle,
  RefreshCw,
  Sparkles,
  UsersRound,
  WifiOff,
  Wrench,
  X,
} from "lucide-react";
import type { Language } from "../stores/editorStore";
import { localizeReleaseNotes } from "../update/releaseNotes";
import type { UpdateCheckResult } from "../update/updateService";

const QQ_GROUP_NUMBER = "1045801770";
const QQ_GROUP_IMAGE_URL = "/qq-group-card.svg";

interface Props {
  open: boolean;
  language: Language;
  checking: boolean;
  error: string;
  result: UpdateCheckResult | null;
  checkOnStartup: boolean;
  automaticPrompt: boolean;
  onCheckOnStartupChange: (enabled: boolean) => void;
  onRetry: () => void;
  onOpenRelease: () => void;
  onClose: () => void;
}

export function UpdateDialog({
  open,
  language,
  checking,
  error,
  result,
  checkOnStartup,
  automaticPrompt,
  onCheckOnStartupChange,
  onRetry,
  onOpenRelease,
  onClose,
}: Props) {
  const dialogRef = useRef<HTMLElement>(null);
  const previousFocusRef = useRef<HTMLElement | null>(null);
  const isEn = language === "en";
  const releaseNotes = useMemo(
    () => localizeReleaseNotes(result?.releaseNotes ?? "", language),
    [language, result?.releaseNotes],
  );

  useEffect(() => {
    if (!open) return;
    previousFocusRef.current = document.activeElement as HTMLElement | null;
    const frame = window.requestAnimationFrame(() => {
      dialogRef.current?.querySelector<HTMLElement>("button")?.focus();
    });
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.preventDefault();
        onClose();
        return;
      }
      if (event.key !== "Tab" || !dialogRef.current) return;
      const focusable = Array.from(
        dialogRef.current.querySelectorAll<HTMLElement>(
          'button:not(:disabled), input:not(:disabled), [tabindex]:not([tabindex="-1"])',
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
  }, [open, onClose]);

  if (!open) return null;

  const updateAvailable = Boolean(result?.updateAvailable);
  const hasReleaseNotes =
    releaseNotes.features.length > 0 ||
    releaseNotes.fixes.length > 0 ||
    releaseNotes.other.length > 0;
  const publishedDate = result?.publishedAt
    ? new Intl.DateTimeFormat(isEn ? "en-US" : "vi-VN", {
        year: "numeric",
        month: "short",
        day: "numeric",
      }).format(new Date(result.publishedAt))
    : "";
  const title = checking
    ? isEn
      ? "Checking for updates"
      : "Đang kiểm tra cập nhật"
    : error
      ? isEn
        ? "Unable to check"
        : "Không thể kiểm tra được"
      : updateAvailable
        ? isEn
          ? "A new version is available"
          : "Đã có phiên bản mới"
        : isEn
          ? "VisualTeX is up to date"
          : "VisualTeX đã được cập nhật";

  return (
    <div className="modal-backdrop update-backdrop" role="presentation">
      <section
        ref={dialogRef}
        className="update-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="update-dialog-title"
      >
        <header className="dialog-header update-dialog-header">
          <div className="update-title-group">
            <span
              className={
                "update-dialog-icon " +
                (error ? "is-error" : updateAvailable ? "is-available" : "")
              }
            >
              {checking ? (
                <LoaderCircle size={19} className="is-spinning" />
              ) : error ? (
                <WifiOff size={19} />
              ) : updateAvailable ? (
                <Download size={19} />
              ) : (
                <CheckCircle2 size={19} />
              )}
            </span>
            <div>
              <h2 id="update-dialog-title">{title}</h2>
              <span>{isEn ? "Application update" : "Cập nhật ứng dụng"}</span>
            </div>
          </div>
          <button
            type="button"
            className="icon-button"
            onClick={onClose}
            aria-label={isEn ? "Close update dialog" : "Đóng hộp thoại cập nhật"}
          >
            <X size={18} />
          </button>
        </header>

        <div className="update-dialog-content">
          {checking ? (
            <p>
              {isEn
                ? "Connecting to the VisualTeX release server…"
                : "Đang kết nối với máy chủ phát hành VisualTeX…"}
            </p>
          ) : error ? (
            <>
              <p>
                {isEn
                  ? "Check your network connection and try again."
                  : "Hãy kiểm tra kết nối mạng của bạn và thử lại."}
              </p>
              <code>{error}</code>
            </>
          ) : result ? (
            <>
              <div className="update-version-row">
                <span>
                  <small>{isEn ? "Installed" : "Đã cài đặt"}</small>
                  <strong>v{result.currentVersion}</strong>
                </span>
                <RefreshCw size={16} aria-hidden="true" />
                <span>
                  <small>{isEn ? "Latest" : "Mới nhất"}</small>
                  <strong>v{result.latestVersion}</strong>
                </span>
              </div>

              {updateAvailable ? (
                <>
                  <div className="update-release-heading">
                    <strong>{result.releaseName}</strong>
                    {publishedDate && (
                      <small>
                        {isEn ? `Published ${publishedDate}` : `Đã xuất bản ${publishedDate}`}
                      </small>
                    )}
                  </div>

                  {hasReleaseNotes ? (
                    <div className="update-release-notes">
                      {releaseNotes.features.length > 0 && (
                        <section>
                          <h3>
                            <Sparkles size={15} />
                            {isEn ? "New features" : "Tính năng mới"}
                          </h3>
                          <ul>
                            {releaseNotes.features.map((item, index) => (
                              <li key={`feature-${index}`}>{item}</li>
                            ))}
                          </ul>
                        </section>
                      )}
                      {releaseNotes.fixes.length > 0 && (
                        <section>
                          <h3>
                            <Wrench size={15} />
                            {isEn ? "Bug fixes" : "Sửa lỗi"}
                          </h3>
                          <ul>
                            {releaseNotes.fixes.map((item, index) => (
                              <li key={`fix-${index}`}>{item}</li>
                            ))}
                          </ul>
                        </section>
                      )}
                      {releaseNotes.other.length > 0 && (
                        <section>
                          <h3>{isEn ? "Other changes" : "Những thay đổi khác"}</h3>
                          <ul>
                            {releaseNotes.other.map((item, index) => (
                              <li key={`other-${index}`}>{item}</li>
                            ))}
                          </ul>
                        </section>
                      )}
                    </div>
                  ) : (
                    <p>
                      {isEn
                        ? "Open GitHub Releases to view the complete update details and download the installer for your platform."
                        : "Mở Bản phát hành GitHub để xem chi tiết cập nhật đầy đủ và tải xuống trình cài đặt cho nền tảng của bạn."}
                    </p>
                  )}

                  <section
                    className="update-community-card"
                    aria-label="VisualTeX QQ communication group"
                  >
                    <div className="update-community-copy">
                      <span className="update-community-icon" aria-hidden="true">
                        <UsersRound size={18} />
                      </span>
                      <div>
                        <strong>Join the VisualTeX QQ communication group</strong>
                        <p>
                          "If you are interested in the use, development or improvement of VisualTeX, you are welcome to scan the QR code or search the group number to join the exchange."</p>
                        <span className="update-community-number">
                          "QQ group number:"<b>{QQ_GROUP_NUMBER}</b>
                        </span>
                      </div>
                    </div>
                    <img
                      src={QQ_GROUP_IMAGE_URL}
                      alt={`VisualTeX QQ group${QQ_GROUP_NUMBER}QR code`}
                      loading="lazy"
                    />
                  </section>
                </>
              ) : (
                <p>
                  {isEn
                    ? "You are using the latest stable version."
                    : "Bạn đang sử dụng phiên bản ổn định mới nhất."}
                </p>
              )}
            </>
          ) : null}

          <label className="update-preference-row">
            <input
              type="checkbox"
              checked={automaticPrompt ? !checkOnStartup : checkOnStartup}
              onChange={(event) =>
                onCheckOnStartupChange(
                  automaticPrompt ? !event.target.checked : event.target.checked,
                )
              }
            />
            <span>
              <strong>
                {automaticPrompt
                  ? isEn
                    ? "Do not remind me again"
                    : "Đừng nhắc nữa"
                  : isEn
                    ? "Check automatically on startup"
                    : "Kiểm tra tự động khi khởi động"}
              </strong>
            </span>
          </label>
        </div>

        <footer className="dialog-footer update-dialog-footer">
          <button type="button" className="secondary-button" onClick={onClose}>
            {isEn ? "Later" : "Sau đó"}
          </button>
          {checking ? (
            <button type="button" className="primary-button" disabled>
              <LoaderCircle size={15} className="is-spinning" />
              {isEn ? "Checking…" : "Đang kiểm tra…"}
            </button>
          ) : error ? (
            <button type="button" className="primary-button" onClick={onRetry}>
              <RefreshCw size={15} /> {isEn ? "Try again" : "Thử lại"}
            </button>
          ) : updateAvailable ? (
            <button type="button" className="primary-button" onClick={onOpenRelease}>
              <Download size={15} /> {isEn ? "Open download page" : "Mở trang tải xuống"}
            </button>
          ) : (
            <button type="button" className="primary-button" onClick={onClose}>
              {isEn ? "Done" : "Xong"}
            </button>
          )}
        </footer>
      </section>
    </div>
  );
}
