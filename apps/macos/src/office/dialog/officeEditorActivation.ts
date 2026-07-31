import type { OfficeHost } from "../api/sessionClient";

export const OFFICE_EDITOR_ACTIVATE_EVENT =
  "visualtex-office-editor-activate";
export const OFFICE_EDITOR_CLEAR_EVENT = "visualtex-office-editor-clear";

export interface OfficeEditorActivation {
  sessionId: string;
  host: OfficeHost;
  generation: number;
  receivedEpochMs: number;
}

export interface OfficeEditorClear {
  sessionId: string;
  generation: number;
}

export function isOfficeEditorActivation(
  value: unknown,
): value is OfficeEditorActivation {
  if (!value || typeof value !== "object") return false;
  const candidate = value as Partial<OfficeEditorActivation>;
  return (
    typeof candidate.sessionId === "string" &&
    /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(
      candidate.sessionId,
    ) &&
    (candidate.host === "word" || candidate.host === "powerpoint") &&
    typeof candidate.generation === "number" &&
    Number.isSafeInteger(candidate.generation) &&
    candidate.generation > 0 &&
    typeof candidate.receivedEpochMs === "number" &&
    Number.isSafeInteger(candidate.receivedEpochMs) &&
    candidate.receivedEpochMs > 0
  );
}

export function shouldAcceptOfficeEditorActivation(
  current: OfficeEditorActivation | null,
  next: OfficeEditorActivation,
  expectedHost: OfficeHost | null,
) {
  return (
    (!expectedHost || next.host === expectedHost) &&
    (!current || next.generation > current.generation)
  );
}

export function clearsOfficeEditorActivation(
  current: OfficeEditorActivation | null,
  payload: OfficeEditorClear,
) {
  return Boolean(
    current &&
      current.sessionId === payload.sessionId &&
      current.generation === payload.generation,
  );
}
