import { useEffect, useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import { listen } from "@tauri-apps/api/event";
import { AlertCircle, CheckCircle2, LoaderCircle } from "lucide-react";
import DesktopShell from "../App";

interface SilentOcrHudPayload {
  status: "running" | "success" | "error";
  message: string;
  progress: number;
}

function SilentOcrHud() {
  const [payload, setPayload] = useState<SilentOcrHudPayload>({
    status: "running",
    message: "Processing silent OCR…",
    progress: 8,
  });

  useEffect(() => {
    document.documentElement.dataset.visualtexView = "silent-ocr-hud";
    let disposed = false;
    let unlisten: (() => void) | undefined;
    const refresh = () => {
      void invoke<SilentOcrHudPayload>("get_silent_ocr_hud_status")
        .then((current) => {
          if (!disposed) setPayload(current);
        })
        .catch(() => undefined);
    };
    refresh();
    const pollTimer = window.setInterval(refresh, 120);
    void listen<SilentOcrHudPayload>("visualtex-silent-ocr-status", (event) => {
      if (!disposed) setPayload(event.payload);
    }).then((dispose) => {
      unlisten = dispose;
    });
    return () => {
      disposed = true;
      window.clearInterval(pollTimer);
      unlisten?.();
      delete document.documentElement.dataset.visualtexView;
    };
  }, []);

  const statusClass =
    payload.status === "running" ? "is-busy" : `is-${payload.status}`;
  const StatusIcon =
    payload.status === "success"
      ? CheckCircle2
      : payload.status === "error"
        ? AlertCircle
        : LoaderCircle;
  const statusTitle =
    payload.status === "success"
      ? "Recognition successful"
      : payload.status === "error"
        ? "Recognition failed"
        : "Silent OCR";
  return (
    <main className="silent-ocr-hud-page">
      <div className={`windows-quick-ocr-hud ${statusClass}`} role="status" aria-live="polite">
        <div className="silent-ocr-hud-icon" aria-hidden="true">
          <StatusIcon size={18} />
        </div>
        <div className="silent-ocr-hud-copy">
          <strong>{statusTitle}</strong>
          <div className="silent-ocr-hud-message">{payload.message}</div>
          <div className="silent-ocr-hud-progress" aria-hidden="true">
            <span style={{ width: `${Math.max(4, Math.min(100, payload.progress))}%` }} />
          </div>
        </div>
      </div>
    </main>
  );
}

export function DesktopApp() {
  const view = new URLSearchParams(window.location.search).get("view");
  if (view === "silent-ocr-hud") return <SilentOcrHud />;
  return <DesktopShell />;
}

export default DesktopApp;
