import { ensureCompanionReady } from "../api/companionClient";
import {
  createOfficeSession,
  getOfficeSession,
  updateOfficeSession,
  type OfficeFormulaSession,
  type OfficeSessionMode,
} from "../shared/sessionClient";
import type {
  OfficeHostAdapter,
  OfficeInteractionTarget,
} from "../adapters/OfficeHostAdapter";
import { officeErrorMessage } from "../errors";
import { DialogController } from "./DialogController";
import type { VisualTeXDialogMessage } from "./bridgeMessages";

function sessionHasFormula(lines: Array<{ latex: string }>) {
  return lines.some((line) => line.latex.trim().length > 0);
}

function sessionHasRequiredExport(
  session: OfficeFormulaSession,
  adapter: OfficeHostAdapter,
) {
  if (!session.exportResult) return false;
  return (
    adapter.requiredExportFormat !== "png" ||
    Boolean(session.exportResult.pngBase64?.trim())
  );
}

function delay(milliseconds: number) {
  return new Promise<void>((resolve) => window.setTimeout(resolve, milliseconds));
}

export type OfficeSessionCommitter = (
  session: OfficeFormulaSession,
  adapter: OfficeHostAdapter,
) => Promise<void>;

async function defaultCommitter(
  session: OfficeFormulaSession,
  adapter: OfficeHostAdapter,
) {
  await adapter.applySession(session);
}

function showCommandError(adapter: OfficeHostAdapter, message: string) {
  adapter.showMessage(message);
  try {
    window.alert(`VisualTeX\n\n${message}`);
  } catch {
    // Some Office hosts suppress modal alerts in command pages.
  }
}

export class OfficeBridge {
  private readonly dialog = new DialogController();
  private activeSessionId: string | null = null;
  private commandRunning = false;
  private commandCompleted: (() => void) | null = null;
  private sessionWatchTimer: number | null = null;
  private sessionWatchRunning = false;
  private commitRunning = false;

  constructor(
    private readonly adapter: OfficeHostAdapter,
    private readonly commitWithPlatform: OfficeSessionCommitter = defaultCommitter,
  ) {}

  async run(
    mode: OfficeSessionMode,
    onCommandCompleted?: () => void,
    interactionTarget?: OfficeInteractionTarget,
    options: { silentFailure?: boolean } = {},
  ) {
    if (this.commandRunning || this.dialog.isOpen) {
      this.adapter.showMessage("The VisualTeX editing window has been opened.");
      onCommandCompleted?.();
      return;
    }

    if (interactionTarget) {
      this.adapter.prepareInteractionTarget?.(interactionTarget);
    }

    this.commandRunning = true;
    this.commandCompleted = onCommandCompleted ?? null;
    try {
      this.adapter.showMessage("Connecting to VisualTeX local companion service...");
      await ensureCompanionReady();
      const selection = await this.adapter.readSelection(mode);
      const session = await createOfficeSession({
        mode,
        host: this.adapter.host,
        sourceDocumentId: selection.sourceDocumentId,
        sourceObjectId: selection.sourceObjectId,
        autoCommitOnClose: true,
        ...selection.sessionSeed,
      });
      this.activeSessionId = session.id;
      this.adapter.showMessage("Opening VisualTeX editor...");
      await this.dialog.open(session.id, {
        onMessage: (message) => this.handleDialogMessage(message),
        onClosed: () => this.handleDialogClosed(session.id),
      });
      this.startSessionWatch(session.id);
      // ExecuteFunction events are short-lived Office commands, not ownership
      // tokens for the entire editor Session. PowerPoint for Mac serializes
      // later ribbon commands while the event is pending, which made the edit
      // button appear dead after one invocation. Release it once the dialog is
      // open; the Session watcher and dialog handlers continue independently.
      this.completeOfficeCommand();
      this.adapter.showMessage("The VisualTeX editor is open.");
    } catch (error) {
      if (!options.silentFailure) {
        showCommandError(
          this.adapter,
          officeErrorMessage(error, "Unable to start the VisualTeX Office editor."),
        );
      }
      this.activeSessionId = null;
      this.finishSession();
    } finally {
      this.commandRunning = false;
    }
  }

  async openDesktopApp() {
    try {
      await this.adapter.openDesktopApp();
    } catch (error) {
      showCommandError(
        this.adapter,
        officeErrorMessage(error, "Unable to open VisualTeX.app."),
      );
    }
  }

  private completeOfficeCommand() {
    const completed = this.commandCompleted;
    this.commandCompleted = null;
    try {
      completed?.();
    } catch {
      // Office can invalidate the command event after the host closes.
    }
  }

  private finishSession() {
    this.stopSessionWatch();
    this.completeOfficeCommand();
  }

  private stopSessionWatch() {
    if (this.sessionWatchTimer !== null) {
      window.clearInterval(this.sessionWatchTimer);
      this.sessionWatchTimer = null;
    }
    this.sessionWatchRunning = false;
  }

  private startSessionWatch(sessionId: string) {
    this.stopSessionWatch();
    this.sessionWatchTimer = window.setInterval(() => {
      if (this.sessionWatchRunning || sessionId !== this.activeSessionId) return;
      this.sessionWatchRunning = true;
      void this.checkSessionState(sessionId).finally(() => {
        this.sessionWatchRunning = false;
      });
    }, 150);
  }

  private async checkSessionState(sessionId: string) {
    const session = await getOfficeSession(sessionId);
    if (sessionId !== this.activeSessionId) return;

    if (session.status === "committing") {
      await this.commitSession(sessionId, true);
      return;
    }

    if (session.status === "cancelled" || session.explicitCancel) {
      this.dialog.close();
      this.activeSessionId = null;
      this.adapter.showMessage("Canceled, Office document has not been modified.");
      this.finishSession();
      return;
    }

    if (session.status === "completed") {
      this.dialog.close();
      this.activeSessionId = null;
      this.finishSession();
    }
  }

  private async handleDialogMessage(message: VisualTeXDialogMessage) {
    if (message.sessionId !== this.activeSessionId) return;

    if (message.type === "visualtex-ready") {
      await updateOfficeSession(message.sessionId, { status: "editing" });
      return;
    }

    if (message.type === "visualtex-cancel") {
      await updateOfficeSession(message.sessionId, {
        status: "cancelled",
        explicitCancel: true,
      });
      this.dialog.close();
      this.activeSessionId = null;
      this.adapter.showMessage("Canceled, Office document has not been modified.");
      this.finishSession();
      return;
    }

    if (message.type === "visualtex-error") {
      await updateOfficeSession(message.sessionId, {
        status: "failed",
        error: message.message,
      });
      this.adapter.showMessage(message.message);
      return;
    }

    if (message.type === "visualtex-commit") {
      await this.commitSession(message.sessionId, true);
    }
  }

  private async handleDialogClosed(sessionId: string) {
    if (sessionId !== this.activeSessionId) return;
    this.stopSessionWatch();

    try {
      let session = await getOfficeSession(sessionId);
      for (
        let attempt = 0;
        // The dialog persists SVG immediately, then finishes PNG
        // rasterization in the background. Windows OLE requires that PNG,
        // so allow the final in-flight save to reach the companion before
        // deciding that a directly closed editor cannot be committed.
        attempt < 50 &&
        session.status !== "completed" &&
        session.status !== "cancelled" &&
        (!sessionHasFormula(session.lines) ||
          !sessionHasRequiredExport(session, this.adapter));
        attempt += 1
      ) {
        await delay(100);
        session = await getOfficeSession(sessionId);
      }
      if (
        session.status === "completed" ||
        session.status === "cancelled" ||
        session.explicitCancel
      ) {
        return;
      }

      const shouldAutoCommit =
        session.autoCommitOnClose &&
        sessionHasFormula(session.lines) &&
        sessionHasRequiredExport(session, this.adapter) &&
        (session.mode === "create" || session.dirty);

      if (shouldAutoCommit) {
        await this.commitSession(sessionId, false);
        return;
      }

      if (!sessionHasFormula(session.lines)) {
        await updateOfficeSession(sessionId, {
          status: "cancelled",
          explicitCancel: false,
        });
        this.adapter.showMessage("The empty formula has been canceled and the Office document has not been modified.");
      } else if (!sessionHasRequiredExport(session, this.adapter)) {
        await updateOfficeSession(sessionId, {
          status: "failed",
          error: "The formula export has not been successful and the recovery record has been retained.",
        });
        this.adapter.showMessage("Formula not inserted: Export failed, Session reserved for recovery.");
      }
    } catch (error) {
      this.adapter.showMessage(
        `Unable to handle shutdown event:${officeErrorMessage(
          error,
          "Unable to handle the VisualTeX edit window close event.",
        )}`,
      );
    } finally {
      this.activeSessionId = null;
      this.finishSession();
    }
  }

  private async commitSession(sessionId: string, closeAfterSuccess: boolean) {
    if (this.commitRunning) return;
    this.commitRunning = true;
    try {
      const session = await getOfficeSession(sessionId);
      if (session.status === "cancelled" || session.explicitCancel) {
        if (closeAfterSuccess) this.dialog.close();
        this.activeSessionId = null;
        this.finishSession();
        return;
      }
      if (!sessionHasFormula(session.lines)) {
        await updateOfficeSession(sessionId, { status: "cancelled" });
        this.adapter.showMessage("An empty formula was not inserted into the Office document.");
        if (closeAfterSuccess) this.dialog.close();
        this.activeSessionId = null;
        this.finishSession();
        return;
      }
      if (!session.exportResult) {
        throw new Error("The formula SVG has not been generated and cannot be written to the Office document.");
      }
      if (session.mode === "edit" && !session.dirty) {
        await updateOfficeSession(sessionId, { status: "completed" });
        this.activeSessionId = null;
        if (closeAfterSuccess) this.dialog.close();
        this.adapter.showMessage("The content of the formula has not changed and does not need to be updated.");
        this.finishSession();
        return;
      }

      await updateOfficeSession(sessionId, { status: "committing", error: null });
      await this.commitWithPlatform(session, this.adapter);
      await updateOfficeSession(sessionId, { status: "completed", error: null });
      this.activeSessionId = null;
      if (closeAfterSuccess) this.dialog.close();
      this.adapter.showMessage(
        session.mode === "edit" ? "VisualTeX formulas have been updated." : "VisualTeX formula inserted.",
      );
      this.finishSession();
    } catch (error) {
      const message = officeErrorMessage(error, "Office formula writing failed.");
      await updateOfficeSession(sessionId, {
        status: "failed",
        error: message,
      }).catch(() => undefined);
      this.adapter.showMessage(`${message}Session has been preserved for recovery.`);
    } finally {
      this.commitRunning = false;
    }
  }
}
