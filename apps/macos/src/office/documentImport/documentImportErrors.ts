import { errorMessage } from "../../runtime/errorMessage.ts";

export function documentImportErrorMessage(reason: unknown, fallback: string): string {
  return errorMessage(reason, fallback);
}
