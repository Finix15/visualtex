import { useEffect, useRef, useState } from "react";
import {
  ArrowLeft,
  ArrowRight,
  Check,
  Code2,
  Download,
  FileCode2,
  FileImage,
  FileText,
  FolderOpen,
  Keyboard,
  Menu,
  MousePointerClick,
  PanelLeft,
  Power,
  Presentation,
  RefreshCw,
  ScanLine,
  Settings2,
  ToggleLeft,
  Trash2,
  X,
} from "lucide-react";
import { MathPreview } from "./MathPreview";
import { VisualTeXLogo } from "./VisualTeXLogo";
import { PowerPointAddinGuide } from "./PowerPointAddinGuide";
import type { Language } from "../stores/editorStore";
import type { DesktopPlatform } from "../platform";
import { THEME_DEFINITIONS } from "../themeCustomization";

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
  | "code-format"
  | "export"
  | "input-behavior"
  | "ocr-setup"
  | "paste-image"
  | "mac-word-plugin"
  | "mac-powerpoint-load"
  | "mac-powerpoint-use"
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
      id: "code-format",
      title: isEn ? "Switch the LaTeX code format" : "Chuyển đổi định dạng mã LaTeX",
      description: isEn
        ? "Choose an independent or combined environment from the top bar. The source panel and copied output update immediately."
        : "Chọn môi trường độc lập hoặc kết hợp từ thanh trên cùng. Bảng nguồn và bản sao chép đầu ra cập nhật ngay lập tức.",
    },
    {
      id: "export",
      title: isEn ? "Export the current document" : "Xuất tài liệu hiện tại",
      description: isEn
        ? "Open Export in the top bar to save the current document as Markdown, SVG, or PNG. Choose the destination once and VisualTeX remembers it for later exports."
        : "Mở Xuất ở thanh trên cùng để lưu tài liệu hiện tại dưới dạng Markdown, SVG hoặc PNG. Chọn đích một lần và VisualTeX sẽ ghi nhớ đích đó để xuất sau này.",
    },
    {
      id: "input-behavior",
      title: isEn ? "Customize input behavior" : "Tùy chỉnh hành vi nhập liệu",
      description: isEn
        ? "Use Input behavior to control automatic exits from scripts, accents, and font commands, plus the large command suggestion panels. With font auto-exit off, type multiple characters and press Enter to confirm."
        : "Sử dụng hành vi Nhập liệu để kiểm soát việc tự động thoát khỏi các lệnh tập lệnh, dấu trọng âm và phông chữ, cùng với bảng gợi ý lệnh lớn. Khi tắt tính năng tự động thoát phông chữ, hãy nhập nhiều ký tự và nhấn Enter để xác nhận.",
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
        id: "mac-word-plugin",
        title: isEn ? "Use the native VisualTeX tab in Word" : "Sử dụng tab VisualTeX gốc trong Word",
        description: isEn
          ? "After installing the DOTM and restarting Word, open the VisualTeX tab. Insert picture or native OMML formulas, edit the selected formula, convert an image formula to Word math, and manage numbering or cross-references."
          : "Sau khi cài đặt DOTM và khởi động lại Word, hãy mở tab VisualTeX. Chèn hình ảnh hoặc công thức OMML gốc, chỉnh sửa công thức đã chọn, chuyển đổi công thức hình ảnh thành toán Word và quản lý việc đánh số hoặc tham chiếu chéo.",
      },
      {
        id: "mac-powerpoint-load",
        title: isEn ? "Register the PPAM once in PowerPoint" : "Đăng ký PPAM một lần trong PowerPoint",
        description: isEn
          ? "Open Tools → PowerPoint Add-ins, click +, select the fixed VisualTeX.ppam file, keep VisualTeX checked, and restart PowerPoint. Later VisualTeX updates reuse the same registered path."
          : "Mở Công cụ → Phần bổ trợ PowerPoint, nhấp vào +, chọn tệp VisualTeX.ppam đã sửa, giữ nguyên VisualTeX được chọn và khởi động lại PowerPoint. Các bản cập nhật VisualTeX sau này sẽ sử dụng lại đường dẫn đã đăng ký tương tự.",
      },
      {
        id: "mac-powerpoint-use",
        title: isEn ? "Create and edit formulas in PowerPoint" : "Tạo và chỉnh sửa công thức trong PowerPoint",
        description: isEn
          ? "Open the VisualTeX tab and choose New formula. Select an existing VisualTeX formula and use Edit selected formula or double-click it to reopen the editor; Delete selected formula removes it cleanly."
          : "Mở tab VisualTeX và chọn Công thức mới. Chọn một công thức VisualTeX hiện có và sử dụng Chỉnh sửa công thức đã chọn hoặc bấm đúp vào công thức đó để mở lại trình chỉnh sửa; Xóa công thức đã chọn sẽ loại bỏ nó một cách sạch sẽ.",
      },
    );
  } else if (platform === "windows") {
    steps.push({
      id: "windows-office-manage",
      title: isEn ? "Manage the Windows OLE service" : "Quản lý dịch vụ Windows OLE",
      description: isEn
        ? "When OLE is selected in the installer, setup completes the certificate, catalog, Ribbon cache, and background registration automatically. In Settings → Windows Office integration, stop the current companion, disable startup, or remove the OLE manifest."
        : "Khi OLE được chọn trong trình cài đặt, quá trình thiết lập sẽ tự động hoàn tất chứng chỉ, danh mục, bộ đệm Ribbon và đăng ký nền. Trong Cài đặt → Tích hợp Windows Office, dừng đồng hành hiện tại, tắt khởi động hoặc xóa tệp kê khai OLE.",
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
                      <kbd>⌥1</kbd>
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
                  <article className="is-selected">
                    <div className="onboarding-layout-mini is-standard" aria-hidden="true">
                      <i />
                      <span />
                      <b />
                    </div>
                    <strong>{isEn ? "Standard layout" : "Bố cục chuẩn"}</strong>
                    <small>{isEn ? "Editor with side tools and tiles" : "Trình chỉnh sửa với các công cụ phụ và ô xếp"}</small>
                  </article>
                  <article>
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
                    <strong>{isEn ? "15 preset themes + Custom" : "15 chủ đề cài sẵn + Tùy chỉnh"}</strong>
                  </header>
                  <div className="onboarding-theme-swatches">
                    {THEME_DEFINITIONS.filter((item) => item.id !== "custom").map(
                      (definition) => (
                        <span key={definition.id}>
                          <i>
                            {definition.swatches.map((color) => (
                              <b key={color} style={{ background: color }} />
                            ))}
                          </i>
                          <small>{isEn ? definition.labelEn : definition.labelVi}</small>
                        </span>
                      ),
                    )}
                  </div>
                  <div className="onboarding-theme-sync-note">
                    <Check size={14} />
                    <span>{isEn ? "Office formula editors inherit the same theme" : "Trình soạn thảo công thức Office kế thừa cùng một chủ đề"}</span>
                  </div>
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
                <div className="onboarding-export-trigger">
                  <Download size={17} />
                  <strong>{isEn ? "Export" : "Xuất khẩu"}</strong>
                  <span>⌄</span>
                </div>
                <div className="onboarding-export-formats">
                  <span><FileText size={18} /><strong>Markdown</strong><small>.md</small></span>
                  <span><FileCode2 size={18} /><strong>SVG</strong><small>.svg</small></span>
                  <span><FileImage size={18} /><strong>PNG</strong><small>.png</small></span>
                </div>
                <div className="onboarding-export-path">
                  <span><FolderOpen size={16} /><small>{isEn ? "Export location" : "Vị trí xuất"}</small></span>
                  <strong>{isEn ? "Choose on first export" : "Chọn ở lần xuất đầu tiên"}</strong>
                </div>
              </div>
            )}

            {current.id === "input-behavior" && (
              <div className="onboarding-input-behavior-demo">
                <div className="onboarding-input-behavior-heading">
                  <MousePointerClick size={17} />
                  <strong>{isEn ? "Input behavior" : "Hành vi nhập liệu"}</strong>
                </div>
                <div className="onboarding-input-behavior-options">
                  <span>
                    <div><strong>{isEn ? "Exit superscript after input" : "Thoát chỉ số trên sau khi nhập"}</strong><small>{isEn ? "Return after one character" : "Trở về sau một ký tự"}</small></div>
                    <i className="is-on" />
                  </span>
                  <span>
                    <div><strong>{isEn ? "Exit font command after input" : "Thoát lệnh phông chữ sau khi nhập"}</strong><small>{isEn ? "Off: type multiple characters" : "Tắt: gõ nhiều ký tự"}</small></div>
                    <i />
                  </span>
                  <span>
                    <div><strong>{isEn ? "Structured command suggestions" : "Gợi ý lệnh có cấu trúc"}</strong><small>{isEn ? "Large VisualTeX command panel" : "Bảng lệnh VisualTeX lớn"}</small></div>
                    <i className="is-on" />
                  </span>
                </div>
                <div className="onboarding-input-behavior-example">
                  <code>\\mathbb&#123;AB&#125;</code>
                  <kbd>Enter</kbd>
                  <small>{isEn ? "Confirm and leave the font scope" : "Xác nhận và để lại phạm vi phông chữ"}</small>
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

            {current.id === "mac-word-plugin" && (
              <div className="onboarding-native-ribbon-demo is-word">
                <div className="onboarding-native-ribbon-title">
                  <FileText size={17} />
                  <strong>Microsoft Word</strong>
                  <span>VisualTeX</span>
                </div>
                <div className="onboarding-native-ribbon-tools">
                  <span><b>OMML</b><small>{isEn ? "Inline" : "Nội tuyến"}</small></span>
                  <span><b>OMML</b><small>{isEn ? "Display" : "Hiển thị"}</small></span>
                  <span><Check size={16} /><small>{isEn ? "Edit selected" : "Đã chọn chỉnh sửa"}</small></span>
                  <span><RefreshCw size={16} /><small>{isEn ? "Update numbers" : "Cập nhật số"}</small></span>
                </div>
                <p>{isEn ? "The DOTM loads automatically after Word restarts." : "DOTM tự động tải sau khi Word khởi động lại."}</p>
              </div>
            )}

            {current.id === "mac-powerpoint-load" && (
              <PowerPointAddinGuide language={language} compact />
            )}

            {current.id === "mac-powerpoint-use" && (
              <div className="onboarding-native-ribbon-demo is-powerpoint">
                <div className="onboarding-native-ribbon-title">
                  <Presentation size={17} />
                  <strong>Microsoft PowerPoint</strong>
                  <span>VisualTeX</span>
                </div>
                <div className="onboarding-native-ribbon-tools">
                  <span><Presentation size={17} /><small>{isEn ? "New formula" : "Công thức mới"}</small></span>
                  <span><Check size={16} /><small>{isEn ? "Edit selected" : "Đã chọn chỉnh sửa"}</small></span>
                  <span><Trash2 size={16} /><small>{isEn ? "Delete selected" : "Xóa đã chọn"}</small></span>
                </div>
                <p>{isEn ? "Double-click an existing VisualTeX formula to edit it again." : "Bấm đúp vào công thức VisualTeX hiện có để chỉnh sửa lại."}</p>
              </div>
            )}

            {current.id === "windows-office-manage" && (
              <div className="onboarding-windows-office-demo">
                <span>
                  <Settings2 size={21} />
                  <strong>{isEn ? "Windows Office integration" : "Tích hợp Windows Office"}</strong>
                </span>
                <div>
                  <span><Power size={19} /><strong>{isEn ? "Stop companion now" : "Dừng đồng hành ngay"}</strong></span>
                  <span><ToggleLeft size={19} /><strong>{isEn ? "Disable startup" : "Tắt khởi động"}</strong></span>
                  <span><Trash2 size={19} /><strong>{isEn ? "Remove OLE manifest" : "Xóa bảng kê khai OLE"}</strong></span>
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
