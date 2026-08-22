import type {
  OfficeHostAdapter,
  OfficeSelectionContext,
} from "../adapters/OfficeHostAdapter";
import { revealDesktopApp } from "../api/companionClient";
import {
  commitWindowsOfficeSession,
  type OfficeFormulaSession,
  type OfficeHost,
  type OfficeSessionMode,
} from "../shared/sessionClient";
import type { OfficeSelectionResult } from "../shared/protocol";
import type { VisualTeXFormulaMetadata } from "../shared/formulaMetadata";
import { OfficeIntegrationError } from "../shared/errors";
import { callWindowsOle } from "./WindowsOleClient";

export interface WindowsOleInteractionTarget {
  host: OfficeHost;
  formulaId: string;
  documentId: string | null;
  objectId: string | null;
  metadata: VisualTeXFormulaMetadata;
}

function selectionMethod(host: OfficeHost) {
  return host === "word"
    ? ("word.getSelection" as const)
    : ("powerpoint.getSelection" as const);
}

export class WindowsOleAdapter implements OfficeHostAdapter {
  readonly requiredExportFormat = "png" as const;
  private pendingInteractionTarget: WindowsOleInteractionTarget | null = null;

  constructor(readonly host: OfficeHost) {}

  prepareWindowsInteractionTarget(target: WindowsOleInteractionTarget) {
    this.pendingInteractionTarget = target.host === this.host ? target : null;
  }

  async readSelection(mode: OfficeSessionMode): Promise<OfficeSelectionContext> {
    const capturedTarget = mode === "edit" ? this.pendingInteractionTarget : null;
    this.pendingInteractionTarget = null;
    if (capturedTarget) {
      return {
        sourceDocumentId: capturedTarget.documentId,
        sourceObjectId: capturedTarget.objectId,
        sessionSeed: {
          formulaId: capturedTarget.metadata.formulaId,
          title: capturedTarget.metadata.title,
          lines: capturedTarget.metadata.lines,
          activeLineId: capturedTarget.metadata.lines[0]?.id ?? null,
          codeFormat: capturedTarget.metadata.codeFormat,
          displayMode: capturedTarget.metadata.displayMode,
          numbered: capturedTarget.metadata.numbered ?? false,
          originalMetadata: capturedTarget.metadata,
        },
      };
    }

    const selection = await callWindowsOle<OfficeSelectionResult>(
      selectionMethod(this.host),
      { mode },
    );
    if (selection.readOnly) {
      throw new OfficeIntegrationError(
        "The current Office document is read-only and formulas cannot be inserted or edited.",
        "document_read_only",
      );
    }
    if (mode === "edit" && (!selection.formulaId || !selection.metadata)) {
      throw new OfficeIntegrationError(
        "Please select a formula with VisualTeX metadata before editing.",
        "formula_not_selected",
      );
    }

    const metadata = selection.metadata;
    return {
      sourceDocumentId: selection.documentId,
      sourceObjectId: selection.objectId,
      sessionSeed: metadata
        ? {
            formulaId: metadata.formulaId,
            title: metadata.title,
            lines: metadata.lines,
            activeLineId: metadata.lines[0]?.id ?? null,
            codeFormat: metadata.codeFormat,
            displayMode: metadata.displayMode,
            numbered: metadata.numbered ?? false,
            originalMetadata: metadata,
          }
        : {
            title: "Office Formula",
            displayMode: this.host === "word" ? "inline" : "block",
            numbered: false,
          },
    };
  }

  async applySession(session: OfficeFormulaSession) {
    await commitWindowsOfficeSession(session.id);
  }

  async updateEquationNumbers() {
    if (this.host !== "word") return 0;
    const result = await callWindowsOle<{ updated: number }>(
      "word.updateEquationNumbers",
      {},
    );
    this.showMessage(`VisualTeX updated${result.updated}formula number.`);
    return result.updated;
  }

  async openDesktopApp() {
    await revealDesktopApp();
  }

  showMessage(message: string) {
    const status = document.getElementById("bridge-status");
    if (status) status.textContent = message;
  }
}

export function windowsOfficeHostFromReadyInfo(host: Office.HostType): OfficeHost {
  if (host === Office.HostType.Word) return "word";
  if (host === Office.HostType.PowerPoint) return "powerpoint";
  throw new OfficeIntegrationError(
    "The VisualTeX Windows OLE plug-in supports Word and PowerPoint only.",
    "unsupported_office_host",
  );
}
