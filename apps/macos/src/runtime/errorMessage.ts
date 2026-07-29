const ERROR_DETAIL_KEYS = [
  "message",
  "error",
  "description",
  "details",
  "reason",
  "cause",
] as const;

function nestedErrorMessage(reason: unknown, seen: Set<object>): string {
  if (typeof reason === "string") return reason.trim();
  if (reason === null || reason === undefined) return "";
  if (typeof reason !== "object") return String(reason).trim();
  if (seen.has(reason)) return "";
  seen.add(reason);

  if (reason instanceof Error) {
    if (reason.message.trim()) return reason.message.trim();
    const cause = (reason as Error & { cause?: unknown }).cause;
    const causeMessage = nestedErrorMessage(cause, seen);
    if (causeMessage) return causeMessage;
    if (reason.name.trim()) return reason.name.trim();
  }

  const record = reason as Record<string, unknown>;
  for (const key of ERROR_DETAIL_KEYS) {
    const message = nestedErrorMessage(record[key], seen);
    if (message) return message;
  }

  try {
    const encoded = JSON.stringify(reason, (_key, value) =>
      typeof value === "bigint" ? value.toString() : value,
    );
    if (encoded && encoded !== "{}") return encoded;
  } catch {
    // Cyclic or host-provided objects fall through to the caller's fallback.
  }
  return "";
}

export function errorMessage(reason: unknown, fallback: string): string {
  return nestedErrorMessage(reason, new Set<object>()) || fallback;
}
