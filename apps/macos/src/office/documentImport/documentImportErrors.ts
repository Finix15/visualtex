import { errorMessage } from "../../runtime/errorMessage";

export function documentImportErrorMessage(reason: unknown, fallback: string): string {
  return errorMessage(reason, fallback);
}
