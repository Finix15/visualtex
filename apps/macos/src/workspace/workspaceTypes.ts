import type { ReactNode, RefObject } from "react";
import type { QuickOcrCaptureMode } from "../ocr/quickOcr";
import type {
  MathEditorHandle,
  MathEditorInsertionTarget,
} from "../editor/MathEditor";
import type { ReplaceDocumentEntry } from "../history/historyTypes";
import type { DocumentSnapshot } from "../history/historyTypes";

export type WorkspaceMode =
  | "desktop"
  | "office-create"
  | "office-edit";

export type WorkspaceExportFormat = "markdown" | "svg" | "png";

export interface WorkspaceOcrModelOption {
  id: string;
  labelVi: string;
  labelEn: string;
}

export interface EditorWorkspaceProps {
  mode: WorkspaceMode;

  showFileActions: boolean;
  showUpdateActions: boolean;
  showOfficeActions: boolean;
  showOcrActions: boolean;

  primaryActionLabel?: string;
  officeHeaderLeadingControls?: ReactNode;
  officeHeaderTrailingActions?: ReactNode;
  desktopHeaderControls?: ReactNode;
  keypadMode?: boolean;

  onPrimaryAction?: () => Promise<void>;
  onCancel?: () => Promise<void>;
  onOpenExport?: () => void;

  editorRef: RefObject<MathEditorHandle | null>;
  editorInstanceKey?: string;
  reuseEditorLineSlots?: boolean;
  sidebarOpen: boolean;
  onSidebarOpenChange: (open: boolean) => void;
  onHistoryBusyChange: (busy: boolean) => void;
  onPasteImage?: (
    file: File,
    target: MathEditorInsertionTarget,
  ) => Promise<void>;
  onCopyPng?: () => Promise<void>;
  onCopy: () => Promise<void>;
  onReplaceDocument: (
    snapshot: DocumentSnapshot,
    source: ReplaceDocumentEntry["source"],
  ) => boolean;

  ocrModel?: string;
  ocrModels?: readonly WorkspaceOcrModelOption[];
  ocrBusy?: boolean;
  onOcrModelChange?: (model: string) => void;
  onQuickOcr?: () => void;
  quickOcrCaptureMode?: QuickOcrCaptureMode;
  onQuickOcrCaptureModeChange?: (mode: QuickOcrCaptureMode) => void;
  silentOcrEnabled?: boolean;
  onSilentOcrEnabledChange?: (enabled: boolean) => void;
  ocrOverlay?: ReactNode;
}
