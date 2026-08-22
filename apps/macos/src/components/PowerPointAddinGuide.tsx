import { Check, FolderOpen, Plus } from "lucide-react";
import type { Language } from "../stores/editorStore";

interface Props {
  language: Language;
  compact?: boolean;
  loaded?: boolean;
}

export function PowerPointAddinGuide({ language, compact = false, loaded = false }: Props) {
  const isEn = language === "en";

  return (
    <div className={`powerpoint-native-guide${compact ? " is-compact" : ""}`}>
      <div className="powerpoint-native-app">
        <div className="powerpoint-native-titlebar">
          <span className="powerpoint-window-controls" aria-hidden="true">
            <i />
            <i />
            <i />
          </span>
          <strong>Microsoft PowerPoint</strong>
        </div>
        <div className="powerpoint-native-menubar">
          <span>PowerPoint</span>
          <span>{isEn ? "File" : "Tập tin"}</span>
          <span>{isEn ? "Edit" : "Chỉnh sửa"}</span>
          <span>{isEn ? "View" : "Xem"}</span>
          <span>{isEn ? "Insert" : "Chèn"}</span>
          <span className="is-active">{isEn ? "Tools" : "Dụng cụ"}</span>
        </div>
        <div className="powerpoint-native-tools-menu">
          <span>{isEn ? "Spelling…" : "Chính tả…"}</span>
          <span>{isEn ? "Language…" : "Ngôn ngữ…"}</span>
          <span className="is-highlighted">
            {isEn ? "PowerPoint Add-ins…" : "Phần bổ trợ PowerPoint…"}
          </span>
        </div>
      </div>

      <div className="powerpoint-addins-dialog-mock">
        <header>
          <strong>{isEn ? "PowerPoint Add-ins" : "Phần bổ trợ PowerPoint"}</strong>
        </header>
        <div className={`powerpoint-addins-list${loaded ? " is-loaded" : " is-empty"}`}>
          {loaded ? (
            <span className="is-selected">
              <i><Check size={12} /></i>
              <strong>VisualTeX</strong>
            </span>
          ) : (
            <span className="powerpoint-addins-empty-copy">
              <strong>{isEn ? "VisualTeX is not listed yet" : "VisualTeX chưa được liệt kê"}</strong>
            </span>
          )}
        </div>
        <div className="powerpoint-addins-controls">
          <span className={loaded ? "" : "is-next-action"} aria-hidden="true"><Plus size={13} /></span>
          <span aria-hidden="true">−</span>

        </div>
      </div>

      <div className="powerpoint-guide-steps">
        <span><b>1</b>{isEn ? "Open Tools → PowerPoint Add-ins" : "Mở Công cụ → Bổ trợ PowerPoint"}</span>
        <span><b>2</b><Plus size={13} />{isEn ? "Click +. VisualTeX will not be in the list before this step" : "Bấm +. VisualTeX sẽ không có trong danh sách trước bước này"}</span>
        <span><b>3</b><Check size={13} />{isEn ? "Keep VisualTeX checked, then restart PowerPoint" : "Luôn kiểm tra VisualTeX, sau đó khởi động lại PowerPoint"}</span>
        <span className="powerpoint-guide-path"><FolderOpen size={13} />~/Library/Group Containers/UBF8T346G9.Office/VisualTeX/OfficeAddins/VisualTeX.ppam</span>
      </div>
    </div>
  );
}
