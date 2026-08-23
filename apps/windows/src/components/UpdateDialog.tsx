import { useEffect, useMemo, useRef } from "react";
import {
  CheckCircle2,
  Download,
  Github,
  LoaderCircle,
  RefreshCw,
  Sparkles,
  Star,
  UserRound,
  UsersRound,
  WifiOff,
  Wrench,
  X,
} from "lucide-react";
import {
  VISUALTEX_QQ_GROUP_NUMBER,
  VISUALTEX_QQ_GROUP_QR_DATA_URL,
} from "../assets/visualtexQqGroup";
import type { Language } from "../stores/editorStore";
import { localizeReleaseNotes } from "../update/releaseNotes";
import type { UpdateCheckResult } from "../update/updateService";

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
  onOpenProject: () => void;
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
  onOpenProject,
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

          <section className="update-project-card" aria-label={isEn ? "Project information" : "Thông tin dự án"}>
            <div className="update-project-author">
              <UserRound size={15} aria-hidden="true" />
              <span>
                <small>{isEn ? "Author" : "Tác giả"}</small>
                <strong>Finix15</strong>
              </span>
            </div>
            <button
              type="button"
              className="update-project-link"
              onClick={onOpenProject}
              title="https://github.com/Finix15/visualtex"
            >
              <Github size={15} aria-hidden="true" />
              <span>github.com/Finix15/visualtex</span>
            </button>
            <p>
              <Star size={14} aria-hidden="true" />
              {isEn
                ? "If you like the project, please give it a Star!"
                : "Nếu bạn thích dự án, hãy cho nó một Ngôi sao!"}
            </p>
          </section>

          <section
            className="update-community-card"
            aria-label={isEn ? "VisualTeX QQ community" : "Cộng đồng VisualTeX QQ"}
          >
            <div className="update-community-copy">
              <span className="update-community-icon">
                <UsersRound size={18} aria-hidden="true" />
              </span>
              <div>
                <small>{isEn ? "Community" : "Cộng đồng"}</small>
                <strong>{isEn ? "VisualTeX QQ Group" : "Nhóm VisualTeX QQ"}</strong>
                <p>
                  {isEn
                    ? "Scan with QQ or search the group number to discuss usage, report issues, and follow development updates."
                    : "Quét bằng QQ hoặc tìm kiếm số nhóm để thảo luận về cách sử dụng, báo cáo sự cố và theo dõi các cập nhật phát triển."}
                </p>
                <span className="update-community-number">
                  {isEn ? "Group number" : "Số nhóm"}：
                  <b>{VISUALTEX_QQ_GROUP_NUMBER}</b>
                </span>
              </div>
            </div>
            <figure className="update-community-qr">
              <img
                src={VISUALTEX_QQ_GROUP_QR_DATA_URL}
                alt={
                  isEn
                    ? `QR code for VisualTeX QQ group ${VISUALTEX_QQ_GROUP_NUMBER}`
                    : `Mã QR cho nhóm VisualTeX QQ ${VISUALTEX_QQ_GROUP_NUMBER}`
                }
                width={240}
                height={240}
              />
              <figcaption>{isEn ? "Scan with QQ" : "Quét bằng QQ"}</figcaption>
            </figure>
          </section>

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
              <small>
                {automaticPrompt
                  ? isEn
                    ? "Automatic update notifications will stay off. You can turn them back on in Settings."
                    : "Thông báo cập nhật tự động sẽ tắt. Bạn có thể bật lại chúng trong Cài đặt."
                  : isEn
                    ? "When disabled, VisualTeX will not make automatic update requests or show update notifications."
                    : "Khi bị tắt, VisualTeX sẽ không thực hiện yêu cầu cập nhật tự động hoặc hiển thị thông báo cập nhật."}
              </small>
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
