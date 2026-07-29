export function documentImportErrorMessage(reason: unknown, fallback: string): string {
  if (reason instanceof Error && reason.message.trim()) return reason.message;
  if (typeof reason === "string" && reason.trim()) return reason;
  if (reason && typeof reason === "object") {
    for (const key of ["message", "error", "description", "details"] as const) {
      const value = (reason as Record<string, unknown>)[key];
      if (typeof value === "string" && value.trim()) return value;
      if (value && typeof value === "object") {
        const nested = documentImportErrorMessage(value, "");
        if (nested) return nested;
      }
    }
    try {
      const encoded = JSON.stringify(reason);
      if (encoded && encoded !== "{}") return encoded;
    } catch {
      // Fall through to the user-facing fallback below.
    }
  }
  return fallback;
}
