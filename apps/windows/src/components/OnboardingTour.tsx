import { useEffect, useRef, useState } from "react";
import {
  ArrowLeft,
  ArrowRight,
  Check,
  Code2,
  Download,
  FileText,
  Grid3X3,
  Keyboard,
  Menu,
  PanelLeft,
  Presentation,
  Puzzle,
  RefreshCw,
  ScanLine,
  Settings2,
  Subscript,
  Superscript,
  ToggleLeft,
  Trash2,
  Type,
  X,
} from "lucide-react";
import { MathPreview } from "./MathPreview";
import { VisualTeXLogo } from "./VisualTeXLogo";
import type { Language } from "../stores/editorStore";
import type { DesktopPlatform } from "../platform";

interface Props {
  open: boolean;
  language: Language;
  platform: DesktopPlatform;
  onFinish: () => void;
}

type StepId =
  | "welcome"
  | "library"
  | "keyboard"
  | "hotkeys-tiles"
  | "layouts-themes"
  | "matrix-fonts"
  | "input-behavior"
  | "code-format"
  | "export"
  | "ocr-setup"
  | "paste-image"
  | "mac-office-enable"
  | "mac-office-manage"
  | "windows-office-manage"
  | "updates";

interface TutorialStep {
  id: StepId;
  title: string;
  description: string;
}

export function tutorialSteps(language: Language, platform: DesktopPlatform): TutorialStep[] {
  const isEn = language === "en";
  const steps: TutorialStep[] = [
    {
      id: "welcome",
      title: isEn ? "Welcome to VisualTeX" : "Chào mừng đến với VisualTeX",
      description: isEn
        ? "Write formulas naturally and inspect the source whenever you need it."
        : "Viết công thức một cách tự nhiên và kiểm tra nguồn bất cứ khi nào bạn cần.",
    },
    {
      id: "library",
      title: isEn ? "Start from the formula library" : "Bắt đầu từ thư viện công thức",
      description: isEn
        ? "Choose a structure or symbol to insert it at the cursor."
        : "Chọn cấu trúc hoặc ký hiệu để chèn vào con trỏ.",
    },
    {
      id: "keyboard",
      title: isEn ? "Keep your hands on the keyboard" : "Giữ tay trên bàn phím",
      description: isEn
        ? "A few keys cover line creation, navigation, and deletion."
        : "Một số phím bao gồm việc tạo, điều hướng và xóa dòng.",
    },
    {
      id: "hotkeys-tiles",
      title: isEn ? "Turn formulas into fast shortcuts" : "Biến công thức thành phím tắt nhanh",
      description: isEn
        ? "Right-click any formula tool or tile to bind a hotkey. Save the current formula as a custom tile, then organize it by section, colour, and shortcut."
        : "Nhấp chuột phải vào bất kỳ công cụ hoặc ô công thức nào để liên kết phím nóng. Lưu công thức hiện tại dưới dạng ô tùy chỉnh, sau đó sắp xếp công thức theo phần, màu sắc và phím tắt.",
    },
    {
      id: "layouts-themes",
      title: isEn ? "Choose your layout and colour theme" : "Chọn bố cục và chủ đề màu sắc của bạn",
      description: isEn
        ? "Open Settings → Appearance & editing to switch between Standard and Classic layouts and five complete colour themes. Office formula editors follow the same theme."
        : "Mở Cài đặt → Giao diện & chỉnh sửa để chuyển đổi giữa bố cục Tiêu chuẩn và Cổ điển cũng như năm chủ đề màu hoàn chỉnh. Các trình soạn thảo công thức văn phòng cũng tuân theo chủ đề tương tự.",
    },
    {
      id: "matrix-fonts",
      title: isEn ? "Build matrices and styled symbols" : "Xây dựng ma trận và ký hiệu theo kiểu",
      description: isEn
        ? "Choose matrix dimensions up to 10 × 10, then insert blackboard bold, calligraphic, Fraktur, bold, and other styled symbols from the formula tools."
        : "Chọn kích thước ma trận lên tới 10 × 10, sau đó chèn các ký hiệu in đậm, thư pháp, Fraktur, in đậm và các ký hiệu được tạo kiểu khác vào bảng đen từ các công cụ công thức.",
    },
    {
      id: "input-behavior",
      title: isEn ? "Control each input scope independently" : "Kiểm soát độc lập từng phạm vi đầu vào",
      description: isEn
        ? "Superscript and subscript auto-exit are independent. Styled-font input can exit after one character, or stay open and grow until Enter. Differential d is normalized upright in derivative and integral contexts."
        : "Tự động thoát chỉ số trên và chỉ số dưới độc lập. Kiểu nhập phông chữ theo kiểu có thể thoát sau một ký tự hoặc vẫn mở và phát triển cho đến khi Enter. Vi phân d được chuẩn hóa thẳng đứng trong bối cảnh đạo hàm và tích phân.",
    },
    {
      id: "code-format",
      title: isEn ? "Switch the LaTeX code format" : "Chuyển đổi định dạng mã LaTeX",
      description: isEn
        ? "Choose an independent or combined environment from the top bar. The source panel and copied output update immediately."
        : "Chọn môi trường độc lập hoặc kết hợp từ thanh trên cùng. Bảng nguồn và bản sao chép đầu ra cập nhật ngay lập tức.",
    },
    {
      id: "export",
      title: isEn ? "Export from one place" : "Xuất từ một nơi",
      description: isEn
        ? "Use Export to save Markdown, SVG, or PNG. Choose the format, edit the file name, and select any permitted destination path."
        : "Sử dụng Xuất để lưu Markdown, SVG hoặc PNG. Chọn định dạng, chỉnh sửa tên tệp và chọn bất kỳ đường dẫn đích nào được phép.",
    },
    {
      id: "ocr-setup",
      title: isEn ? "First-time OCR setup" : "Thiết lập OCR lần đầu",
      description:
        platform === "macos"
          ? isEn
            ? "The complete macOS package includes Python, PaddleOCR, and the default M model. Setup verifies and extracts the local archives."
            : "Gói macOS hoàn chỉnh bao gồm Python, PaddleOCR và mô hình M mặc định. Thiết lập xác minh và trích xuất các kho lưu trữ cục bộ."
          : platform === "windows"
            ? isEn
              ? "The Windows installer checks for a compatible 64-bit Python 3.9–3.13 runtime. OCR setup is available after that prerequisite is present."
              : "Trình cài đặt Windows kiểm tra thời gian chạy Python 3.9–3.13 64-bit tương thích. Thiết lập OCR khả dụng sau khi có điều kiện tiên quyết đó."
            : isEn
              ? "Open Formula image OCR from the app menu and follow the local runtime setup."
              : "Mở OCR hình ảnh Công thức từ menu ứng dụng và làm theo thiết lập thời gian chạy cục bộ.",
    },
    {
      id: "paste-image",
      title: isEn ? "Paste images directly afterward" : "Dán hình ảnh trực tiếp sau đó",
      description: isEn
        ? "Once OCR is ready, paste an image into a formula field and the result returns to the saved cursor."
        : "Sau khi OCR sẵn sàng, hãy dán hình ảnh vào trường công thức và kết quả sẽ quay trở lại con trỏ đã lưu.",
    },
  ];

  if (platform === "macos") {
    steps.push(
      {
        id: "mac-office-enable",
        title: isEn ? "Enable VisualTeX in Office" : "Kích hoạt VisualTeX trong Office",
        description: isEn
          ? "In Word or PowerPoint, open Home → Add-ins → My Add-ins or Developer Add-ins, then choose VisualTeX. Repeat this after a restart if Office hides the sideloaded tab."
          : "Trong Word hoặc PowerPoint, mở Trang chủ → Phần bổ trợ → Phần bổ trợ của tôi hoặc Phần bổ trợ dành cho nhà phát triển, sau đó chọn VisualTeX. Lặp lại điều này sau khi khởi động lại nếu Office ẩn tab đã tải.",
      },
      {
        id: "mac-office-manage",
        title: isEn ? "Manage the macOS integration" : "Quản lý tích hợp macOS",
        description: isEn
          ? "Open Settings → macOS Office integration. Disable startup without removing the add-in, stop the current companion, or choose Uninstall Office integration to remove it."
          : "Mở Cài đặt → Tích hợp macOS Office. Tắt tính năng khởi động mà không xóa phần bổ trợ, dừng phần bổ trợ hiện tại hoặc chọn Gỡ cài đặt tích hợp Office để xóa phần bổ trợ đó.",
      },
    );
  } else if (platform === "windows") {
    steps.push({
      id: "windows-office-manage",
      title: isEn ? "Use VisualTeX in Word and PowerPoint" : "Sử dụng VisualTeX trong Word và PowerPoint",
      description: isEn
        ? "The native Office add-in lets Word insert inline or display formulas as editable OLE or native OMML, convert formats, update equation numbers, and insert references. In PowerPoint, create or edit formulas, convert them to native OLE, or export them as pictures. Formulas stay with the document and can be reopened by double-clicking."
        : "Phần bổ trợ Office gốc cho phép Word chèn nội tuyến hoặc hiển thị các công thức dưới dạng OLE có thể chỉnh sửa hoặc OMML gốc, chuyển đổi định dạng, cập nhật số phương trình và chèn tham chiếu. Trong PowerPoint, tạo hoặc chỉnh sửa công thức, chuyển đổi chúng thành OLE gốc hoặc xuất chúng dưới dạng ảnh. Các công thức vẫn còn trong tài liệu và có thể được mở lại bằng cách bấm đúp.",
    });
  }

  steps.push({
    id: "updates",
    title: isEn ? "Check for updates anytime" : "Kiểm tra cập nhật bất cứ lúc nào",
    description: isEn
      ? "Open the top-left menu and choose Check for updates. The same action is also available in Settings."
      : "Mở menu trên cùng bên trái và chọn Kiểm tra cập nhật. Hành động tương tự cũng có sẵn trong Cài đặt.",
  });
  return steps;
}

export function OnboardingTour({ open, language, platform, onFinish }: Props) {
  const [step, setStep] = useState(0);
  const dialogRef = useRef<HTMLElement>(null);
  const isEn = language === "en";
  const steps = tutorialSteps(language, platform);
  const current = steps[Math.min(step, steps.length - 1)];
  const lastStep = step === steps.length - 1;

  useEffect(() => {
    if (!open) return;
    setStep(0);
    const frame = window.requestAnimationFrame(() => {
      dialogRef.current?.querySelector<HTMLElement>("button")?.focus();
    });
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.preventDefault();
        onFinish();
        return;
      }
      if (event.key !== "Tab" || !dialogRef.current) return;
      const focusable = Array.from(
        dialogRef.current.querySelectorAll<HTMLElement>(
          'button:not(:disabled), [tabindex]:not([tabindex="-1"])',
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
    };
  }, [open, onFinish]);

  useEffect(() => {
    if (step >= steps.length) setStep(Math.max(steps.length - 1, 0));
  }, [step, steps.length]);

  if (!open || !current) return null;

  const pasteShortcut = platform === "windows" ? "Ctrl+V" : "⌘V";
  const platformLabel =
    platform === "windows"
      ? isEn ? "Formula workspace for Windows" : "Không gian làm việc công thức dành cho Windows"
      : platform === "macos"
        ? isEn ? "Formula workspace for macOS" : "Không gian làm việc công thức dành cho macOS"
        : isEn ? "Visual formula workspace" : "Không gian làm việc công thức trực quan";

  return (
    <div className="onboarding-backdrop">
      <section
        ref={dialogRef}
        className="onboarding-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="onboarding-title"
      >
        <header className="onboarding-header">
          <div className="onboarding-brand">
            <span><VisualTeXLogo className="onboarding-brand-logo" /></span>
            <strong>VisualTeX</strong>
          </div>
          <button
            type="button"
            className="icon-button compact"
            onClick={onFinish}
            aria-label={isEn ? "Close tutorial" : "Đóng hướng dẫn"}
          >
            <X size={16} />
          </button>
        </header>

        <div className="onboarding-content" aria-live="polite">
          <div className="onboarding-copy">
            <span>{String(step + 1).padStart(2, "0")}</span>
            <h2 id="onboarding-title">{current.title}</h2>
            <p>{current.description}</p>
          </div>

          <div className={`onboarding-stage step-${current.id}`}>
            {current.id === "welcome" && (
              <div className="onboarding-welcome-mark">
                <span><VisualTeXLogo className="onboarding-welcome-logo" /></span>
                <div>
                  <strong>VisualTeX</strong>
                  <small>{platformLabel}</small>
                </div>
              </div>
            )}

            {current.id === "library" && (
              <div className="onboarding-library-demo">
                <div className="onboarding-library-rail">
                  <PanelLeft size={15} />
                  <span>{isEn ? "Formula tools" : "Công cụ công thức"}</span>
                </div>
                <div className="onboarding-formula-grid">
                  {["\\frac{a}{b}", "\\sqrt{x}", "\\int_a^b f(x)\\,dx", "\\sum_{i=1}^{n} a_i"].map((latex) => (
                    <span key={latex}><MathPreview latex={latex} /></span>
                  ))}
                </div>
              </div>
            )}

            {current.id === "keyboard" && (
              <div className="onboarding-editor-demo">
                <div className="onboarding-formula-line">
                  <MathPreview latex="\\int_{-\\infty}^{\\infty} e^{-x^2}\\,dx = \\sqrt{\\pi}" />
                </div>
                <div className="onboarding-key-row">
                  <span><Keyboard size={14} /><kbd>Enter</kbd><small>{isEn ? "New line" : "Dòng mới"}</small></span>
                  <span><kbd>Tab</kbd><small>{isEn ? "Next field" : "Trường tiếp theo"}</small></span>
                  <span><kbd>⌫</kbd><small>{isEn ? "Delete empty line" : "Xóa dòng trống"}</small></span>
                </div>
              </div>
            )}

            {current.id === "hotkeys-tiles" && (
              <div className="onboarding-hotkeys-tiles-demo">
                <section className="onboarding-hotkey-guide">
                  <header>
                    <Keyboard size={17} />
                    <strong>{isEn ? "Formula hotkeys" : "Phím nóng công thức"}</strong>
                  </header>
                  <div className="onboarding-hotkey-flow">
                    <span>
                      <b>1</b>
                      <small>{isEn ? "Right-click a tool or tile" : "Nhấp chuột phải vào một công cụ hoặc ô"}</small>
                    </span>
                    <i><ArrowRight size={14} /></i>
                    <span>
                      <b>2</b>
                      <small>{isEn ? "Press and assign a shortcut" : "Nhấn và gán phím tắt"}</small>
                    </span>
                    <i><ArrowRight size={14} /></i>
                    <span>
                      <kbd>Alt+1</kbd>
                      <small>{isEn ? "Insert at the active cursor" : "Chèn vào con trỏ đang hoạt động"}</small>
                    </span>
                  </div>
                  <div className="onboarding-hotkey-manager-note">
                    <Settings2 size={14} />
                    <span>{isEn ? "Settings → Manage formula hotkeys" : "Cài đặt → Quản lý phím nóng công thức"}</span>
                  </div>
                </section>

                <section className="onboarding-custom-tile-guide">
                  <header>
                    <PanelLeft size={17} />
                    <strong>{isEn ? "Custom formula tiles" : "Gạch công thức tùy chỉnh"}</strong>
                  </header>
                  <div className="onboarding-custom-tile-preview">
                    <MathPreview latex="\\int_0^1 x^2\\,\\mathrm{d}x" />
                  </div>
                  <button type="button" tabIndex={-1}>
                    {isEn ? "Save current formula" : "Lưu công thức hiện tại"}
                  </button>
                  <small>
                    {isEn
                      ? "Create sections, then right-click a tile to change its shortcut, colour, or section."
                      : "Tạo các phần, sau đó nhấp chuột phải vào ô để thay đổi phím tắt, màu hoặc phần của ô đó."}
                  </small>
                </section>
              </div>
            )}

            {current.id === "layouts-themes" && (
              <div className="onboarding-layout-theme-demo">
                <div className="onboarding-layout-choice-list">
                  <article>
                    <div className="onboarding-layout-mini is-standard" aria-hidden="true">
                      <i />
                      <span />
                      <b />
                    </div>
                    <strong>{isEn ? "Standard layout" : "Bố cục chuẩn"}</strong>
                    <small>{isEn ? "Editor with side tools and tiles" : "Trình chỉnh sửa với các công cụ phụ và ô xếp"}</small>
                  </article>
                  <article className="is-selected">
                    <div className="onboarding-layout-mini is-classic" aria-hidden="true">
                      <i />
                      <span />
                      <b />
                    </div>
                    <strong>{isEn ? "Classic layout" : "Bố cục cổ điển"}</strong>
                    <small>{isEn ? "Bottom tools with a right tile rail" : "Dụng cụ phía dưới có thanh ray bên phải"}</small>
                  </article>
                </div>

                <div className="onboarding-theme-guide">
                  <header>
                    <Settings2 size={17} />
                    <strong>{isEn ? "Five complete themes" : "Năm chủ đề hoàn chỉnh"}</strong>
                  </header>
                  <div className="onboarding-theme-swatches">
                    {[
                      ["light", isEn ? "Light" : "Ánh sáng"],
                      ["beige", isEn ? "Warm beige" : "Màu be ấm áp"],
                      ["dark", isEn ? "Dark" : "Tối"],
                      ["purple", isEn ? "Deep purple" : "Tím đậm"],
                      ["green", isEn ? "Deep green" : "Xanh đậm"],
                    ].map(([themeId, label]) => (
                      <span className={`is-${themeId}`} key={themeId}>
                        <i><b /><b /><b /></i>
                        <small>{label}</small>
                      </span>
                    ))}
                  </div>
                  <div className="onboarding-theme-sync-note">
                    <Check size={14} />
                    <span>{isEn ? "Office formula editors inherit the same theme" : "Trình soạn thảo công thức Office kế thừa cùng một chủ đề"}</span>
                  </div>
                </div>
              </div>
            )}

            {current.id === "matrix-fonts" && (
              <div className="onboarding-matrix-font-demo">
                <div className="onboarding-matrix-picker-preview">
                  <span className="onboarding-demo-heading">
                    <Grid3X3 size={16} />
                    <strong>{isEn ? "Matrix size" : "Kích thước ma trận"}</strong>
                    <small>10 × 10</small>
                  </span>
                  <div className="onboarding-mini-matrix-grid">
                    {Array.from({ length: 16 }, (_, index) => (
                      <i key={index} className={index < 11 ? "is-selected" : ""} />
                    ))}
                  </div>
                  <b>3 × 4</b>
                </div>
                <i className="onboarding-feature-arrow"><ArrowRight size={15} /></i>
                <div className="onboarding-font-variants-preview">
                  <span className="onboarding-demo-heading">
                    <Type size={16} />
                    <strong>{isEn ? "Font variants" : "Các biến thể phông chữ"}</strong>
                  </span>
                  <div>
                    <span><MathPreview latex="\\mathbb{R}" /><small>mathbb</small></span>
                    <span><MathPreview latex="\\mathcal{F}" /><small>mathcal</small></span>
                    <span><MathPreview latex="\\mathfrak{g}" /><small>mathfrak</small></span>
                  </div>
                </div>
              </div>
            )}

            {current.id === "input-behavior" && (
              <div className="onboarding-input-behavior-demo">
                <div className="onboarding-input-toggle-list">
                  <span>
                    <Superscript size={17} />
                    <strong>{isEn ? "Superscript auto-exit" : "Tự động thoát chỉ số trên"}</strong>
                    <i className="is-on"><b /></i>
                  </span>
                  <span>
                    <Subscript size={17} />
                    <strong>{isEn ? "Subscript auto-exit" : "Tự động thoát đăng ký"}</strong>
                    <i><b /></i>
                  </span>
                </div>
                <div className="onboarding-wrapper-input-preview">
                  <small>{isEn ? "Continuous styled input" : "Đầu vào theo kiểu liên tục"}</small>
                  <span><MathPreview latex="\\mathbb{AB}" /><i /></span>
                  <kbd>Enter</kbd>
                  <ArrowRight size={15} />
                  <MathPreview latex="\\mathbb{AB}C" />
                </div>
                <div className="onboarding-upright-preview">
                  <small>{isEn ? "Automatic upright differential" : "Vi sai thẳng đứng tự động"}</small>
                  <MathPreview latex="\\frac{\\mathrm{d}\\Phi}{\\mathrm{d}\\theta}" />
                </div>
              </div>
            )}

            {current.id === "code-format" && (
              <div className="onboarding-code-format-demo">
                <div className="onboarding-code-format-toolbar">
                  <Code2 size={16} />
                  <strong>{isEn ? "LaTeX code format" : "Định dạng mã LaTeX"}</strong>
                  <span>⌄</span>
                </div>
                <div className="onboarding-code-format-choice">
                  <span>
                    <small>{isEn ? "Independent" : "Độc lập"}</small>
                    <strong>\\[ ... \\]</strong>
                  </span>
                  <span className="is-selected">
                    <Check size={14} />
                    <small>{isEn ? "Combined" : "Kết hợp"}</small>
                    <strong>align*</strong>
                  </span>
                </div>
                <i><ArrowRight size={15} /></i>
                <pre>{"\\begin{align*}\na &= b + c \\\\\\nd &= e - f\n\\end{align*}"}</pre>
              </div>
            )}

            {current.id === "export" && (
              <div className="onboarding-export-demo">
                <div className="onboarding-export-formats">
                  {[
                    ["Markdown", ".md"],
                    ["SVG", ".svg"],
                    ["PNG", ".png"],
                  ].map(([name, extension], index) => (
                    <span key={name} className={index === 1 ? "is-selected" : ""}>
                      <Download size={18} />
                      <strong>{name}</strong>
                      <small>{extension}</small>
                    </span>
                  ))}
                </div>
                <div className="onboarding-export-path">
                  <FileText size={16} />
                  <span>
                    <small>{isEn ? "Save to" : "Lưu vào"}</small>
                    <strong>{isEn ? "Documents / formula.svg" : "Tài liệu / công thức.svg"}</strong>
                  </span>
                  <button type="button" tabIndex={-1}>
                    {isEn ? "Choose…" : "Chọn…"}
                  </button>
                </div>
              </div>
            )}

            {current.id === "ocr-setup" && (
              <div className="onboarding-ocr-setup-demo">
                <span>
                  <ScanLine size={20} />
                  <strong>{isEn ? "Open Formula image OCR" : "Mở OCR hình ảnh công thức"}</strong>
                  <small>{isEn ? "From the app menu" : "Từ menu ứng dụng"}</small>
                </span>
                <i><ArrowRight size={15} /></i>
                <span>
                  <Download size={20} />
                  <strong>{isEn ? "Prepare runtime" : "Chuẩn bị thời gian chạy"}</strong>
                  <small>{platform === "windows" ? "Python 3.9–3.13 x64" : isEn ? "One-time setup" : "Thiết lập một lần"}</small>
                </span>
                <i><ArrowRight size={15} /></i>
                <span>
                  <Check size={20} />
                  <strong>{isEn ? "Verify locally" : "Xác minh cục bộ"}</strong>
                  <small>{isEn ? "Ready for recognition" : "Sẵn sàng để được công nhận"}</small>
                </span>
              </div>
            )}

            {current.id === "paste-image" && (
              <div className="onboarding-paste-demo">
                <div className="onboarding-paste-field">
                  <span className="onboarding-paste-caret" />
                  <small>{isEn ? "Formula field" : "Trường công thức"}</small>
                </div>
                <span className="onboarding-paste-shortcut">
                  <ScanLine size={20} />
                  <strong>{isEn ? "Paste formula image" : "Dán hình ảnh công thức"}</strong>
                  <kbd>{pasteShortcut}</kbd>
                </span>
                <i><ArrowRight size={15} /></i>
                <span className="onboarding-paste-result">
                  <Code2 size={20} />
                  <strong>{isEn ? "Inserted at saved cursor" : "Đã chèn vào con trỏ đã lưu"}</strong>
                  <MathPreview latex="\\frac{a+b}{c}" />
                </span>
              </div>
            )}

            {current.id === "mac-office-enable" && (
              <div className="onboarding-workflow-demo onboarding-office-demo">
                <span>
                  <FileText size={20} />
                  <Presentation size={20} />
                  <strong>Word / PowerPoint</strong>
                </span>
                <i><ArrowRight size={15} /></i>
                <span>
                  <Puzzle size={22} />
                  <strong>{isEn ? "Home → Add-ins" : "Trang chủ → Tiện ích bổ sung"}</strong>
                </span>
                <i><ArrowRight size={15} /></i>
                <span className="is-selected">
                  <Check size={22} />
                  <strong>VisualTeX</strong>
                </span>
              </div>
            )}

            {current.id === "mac-office-manage" && (
              <div className="onboarding-workflow-demo onboarding-office-demo">
                <span>
                  <Settings2 size={22} />
                  <strong>{isEn ? "Settings" : "Cài đặt"}</strong>
                </span>
                <i><ArrowRight size={15} /></i>
                <span>
                  <ToggleLeft size={22} />
                  <strong>{isEn ? "Disable startup" : "Tắt khởi động"}</strong>
                </span>
                <i><ArrowRight size={15} /></i>
                <span className="is-danger">
                  <Trash2 size={22} />
                  <strong>{isEn ? "Uninstall integration" : "Gỡ cài đặt tích hợp"}</strong>
                </span>
              </div>
            )}

            {current.id === "windows-office-manage" && (
              <div className="onboarding-windows-office-demo">
                <span>
                  <FileText size={20} />
                  <Presentation size={20} />
                  <strong>{isEn ? "VisualTeX native Office add-in" : "Bổ trợ Office gốc VisualTeX"}</strong>
                </span>
                <div>
                  <span><Code2 size={19} /><strong>{isEn ? "OLE / OMML · inline / display" : "OLE / OMML · nội tuyến / hiển thị"}</strong></span>
                  <span><RefreshCw size={19} /><strong>{isEn ? "Edit · convert · number · reference" : "Chỉnh sửa · chuyển đổi · số · tham khảo"}</strong></span>
                  <span><Check size={19} /><strong>{isEn ? "Double-click edit · save · export" : "Nhấp đúp vào chỉnh sửa · lưu · xuất"}</strong></span>
                </div>
              </div>
            )}

            {current.id === "updates" && (
              <div className="onboarding-update-demo">
                <span>
                  <Menu size={20} />
                  <strong>{isEn ? "Open app menu" : "Mở menu ứng dụng"}</strong>
                </span>
                <i><ArrowRight size={15} /></i>
                <span className="onboarding-update-menu-item">
                  <RefreshCw size={20} />
                  <strong>{isEn ? "Check for updates" : "Kiểm tra cập nhật"}</strong>
                </span>
                <i><ArrowRight size={15} /></i>
                <span>
                  <Check size={20} />
                  <strong>{isEn ? "Review the result" : "Xem lại kết quả"}</strong>
                </span>
              </div>
            )}
          </div>
        </div>

        <footer className="onboarding-footer">
          <button type="button" className="onboarding-skip" onClick={onFinish}>
            {isEn ? "Skip" : "Bỏ qua"}
          </button>
          <div className="onboarding-progress" aria-label={isEn ? "Tutorial progress" : "Tiến trình hướng dẫn"}>
            {steps.map((item, index) => (
              <span key={item.id} className={index === step ? "is-active" : index < step ? "is-complete" : ""} />
            ))}
          </div>
          <div className="onboarding-actions">
            {step > 0 && (
              <button type="button" className="secondary-button" onClick={() => setStep((value) => value - 1)}>
                <ArrowLeft size={15} />
                {isEn ? "Back" : "Trở lại"}
              </button>
            )}
            <button
              type="button"
              className="primary-button"
              onClick={() => lastStep ? onFinish() : setStep((value) => value + 1)}
            >
              {lastStep ? <Check size={15} /> : null}
              {lastStep ? (isEn ? "Start editing" : "Bắt đầu chỉnh sửa") : (isEn ? "Continue" : "Tiếp tục")}
              {!lastStep ? <ArrowRight size={15} /> : null}
            </button>
          </div>
        </footer>
      </section>
    </div>
  );
}
