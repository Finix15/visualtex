import { useCallback, useEffect, useRef, useState } from "react";
import { readErrorMessage } from "../../errors/readErrorMessage";
import {
  getOfficeSession,
  updateOfficeSession,
  type OfficeFormulaSession,
  type UpdateOfficeSessionInput,
} from "../api/sessionClient";

type OfficeSessionWindow = Window & {
  __VISUALTEX_OFFICE_SESSION_ID__?: string;
};

function sessionIdFromLocation() {
  const injected = (window as OfficeSessionWindow).__VISUALTEX_OFFICE_SESSION_ID__;
  if (injected) return injected;
  const query = new URLSearchParams(window.location.search).get("sessionId");
  if (query) return query;
  const match = window.location.pathname.match(/\/dialog\/([^/?#]+)/);
  return match ? decodeURIComponent(match[1]) : "";
}

export function useOfficeSession() {
  const [sessionId, setSessionId] = useState(sessionIdFromLocation);
  const [session, setSession] = useState<OfficeFormulaSession | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const saveQueueRef = useRef<Promise<void>>(Promise.resolve());

  useEffect(() => {
    const handleSessionChange = (event: Event) => {
      const detail = (event as CustomEvent<{ sessionId?: string }>).detail;
      const next = detail?.sessionId?.trim() ?? "";
      (window as OfficeSessionWindow).__VISUALTEX_OFFICE_SESSION_ID__ = next || undefined;
      saveQueueRef.current = Promise.resolve();
      setSession(null);
      setError("");
      setLoading(Boolean(next));
      setSessionId(next);
    };
    window.addEventListener("visualtex-office-session", handleSessionChange);
    return () => {
      window.removeEventListener("visualtex-office-session", handleSessionChange);
    };
  }, []);

  const reload = useCallback(async () => {
    if (!sessionId) {
      setError("Missing VisualTeX Office session id.");
      setLoading(false);
      return null;
    }
    setLoading(true);
    try {
      const next = await getOfficeSession(sessionId);
      setSession(next);
      setError("");
      return next;
    } catch (reason) {
      setError(readErrorMessage(reason, "Unable to load Office session."));
      return null;
    } finally {
      setLoading(false);
    }
  }, [sessionId]);

  useEffect(() => {
    void reload();
  }, [reload]);

  const save = useCallback(
    (update: UpdateOfficeSessionInput) => {
      if (!sessionId) {
        return Promise.reject(
          new Error("Missing VisualTeX Office session id."),
        );
      }

      // Office autosave and the explicit commit button can fire almost at the
      // same time. Serialize PATCH requests so an older autosave can never
      // arrive after, and overwrite, a committing Session.
      const request = saveQueueRef.current
        .catch(() => undefined)
        .then(() => updateOfficeSession(sessionId, update));
      saveQueueRef.current = request.then(
        () => undefined,
        () => undefined,
      );
      return request.then((next) => {
        setSession(next);
        return next;
      });
    },
    [sessionId],
  );

  return {
    sessionId,
    session,
    loading,
    error,
    reload,
    save,
  };
}
