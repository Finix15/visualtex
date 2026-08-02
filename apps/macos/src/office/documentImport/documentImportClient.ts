import type { VisualTeXFormulaMetadata } from "../shared/formulaMetadata";
import { invokeTauri } from "../shared/tauriTransport";
import type {
  DocumentFormulaDisplayMode,
  DocumentFormulaOutputKind,
  DocumentListKind,
  DocumentParagraphAlignment,
  DocumentParagraphStyle,
} from "./documentImportParser";

export interface MacosDocumentImportRequest {
  protocolVersion: number;
  sessionId: string;
  host: "word";
  sourceDocumentId: string;
  bookmarkName: string;
  defaultFontSizePt: number;
  operation: "documentImport" | "latexRedraw";
  redrawScope?: "selection" | "document";
  outputKind?: DocumentFormulaOutputKind;
  source?: string;
}

export interface DocumentImportParagraphCommitMetadata {
  paragraphId?: string;
  paragraphStyle?: DocumentParagraphStyle;
  paragraphAlignment?: DocumentParagraphAlignment;
  listKind?: DocumentListKind;
  listLevel?: number;
  paragraphStart?: boolean;
  paragraphEnd?: boolean;
}

export interface DocumentImportTextCommitItem
  extends DocumentImportParagraphCommitMetadata {
  kind: "text";
  text: string;
}

export interface DocumentImportFormulaCommitItem
  extends DocumentImportParagraphCommitMetadata {
  kind: "formula";
  formulaId: string;
  latex: string;
  displayMode: DocumentFormulaDisplayMode;
  numbered: boolean;
  fontSizePt: number;
  metadata: VisualTeXFormulaMetadata;
  ommlBase64: string;
  ommlDocxBase64: string;
  svgBase64?: string;
  pngBase64?: string;
  width?: number;
  height?: number;
  baseline?: number;
  sourceStart?: number;
  sourceEnd?: number;
  sourceText?: string;
}

export type DocumentImportCommitItem =
  | DocumentImportTextCommitItem
  | DocumentImportFormulaCommitItem;

export interface CommitMacosDocumentImportInput {
  outputKind: DocumentFormulaOutputKind;
  items: DocumentImportCommitItem[];
}

export interface MacosDocumentImportProgress {
  current: number;
  total: number;
  stage: "preparing" | "inserting" | "complete" | "error" | string;
}

export interface MacosLatexRedrawFontRangeInput {
  sourceStart: number;
  sourceEnd: number;
  sourceText: string;
}

export function getMacosDocumentImportRequest(sessionId: string) {
  return invokeTauri<MacosDocumentImportRequest>(
    "get_macos_offline_document_import_request",
    { sessionId },
  );
}

export function resolveMacosLatexRedrawFontSizes(
  sessionId: string,
  ranges: MacosLatexRedrawFontRangeInput[],
) {
  return invokeTauri<number[]>(
    "resolve_macos_offline_latex_redraw_font_sizes",
    { sessionId, input: { ranges } },
  );
}

export function focusMacosDocumentImportTarget(
  operation: "documentImport" | "latexRedraw" = "documentImport",
) {
  return invokeTauri<void>("focus_macos_offline_document_import_target", {
    operation,
  });
}

export function restoreMacosDocumentImportWindow() {
  return invokeTauri<void>("restore_macos_offline_document_import_window", {});
}

export function getMacosDocumentImportProgress(sessionId: string) {
  return invokeTauri<MacosDocumentImportProgress>(
    "get_macos_offline_document_import_progress",
    { sessionId },
  );
}

export function commitMacosDocumentImport(
  sessionId: string,
  input: CommitMacosDocumentImportInput,
) {
  return invokeTauri<void>("commit_macos_offline_document_import", {
    sessionId,
    input,
  });
}

export function cancelMacosDocumentImport(sessionId: string) {
  return invokeTauri<void>("cancel_macos_offline_document_import", {
    sessionId,
  });
}

export function closeMacosDocumentImportWindow() {
  return invokeTauri<void>("close_macos_offline_office_editor_window", {});
}
