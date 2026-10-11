import { useEffect, useState } from "react";
import { DocxViewer } from "@orim/react";
import type { ChangesMode } from "@orim/core";
import sampleUrl from "../../../fixtures/sample.docx?url";

export function App() {
  const [src, setSrc] = useState<Blob | null>(null);
  const [changes, setChanges] = useState<ChangesMode>("markup");
  const [comments, setComments] = useState(true);
  const [error, setError] = useState("");
  const [dropVisible, setDropVisible] = useState(false);

  useEffect(() => {
    const controller = new AbortController();
    void fetch(sampleUrl, { signal: controller.signal })
      .then((response) => {
        if (!response.ok) throw new Error(`Unable to load sample: ${response.status}`);
        return response.blob();
      })
      .then(setSrc)
      .catch((reason: unknown) => {
        if (!controller.signal.aborted)
          setError(reason instanceof Error ? reason.message : String(reason));
      });
    return () => controller.abort();
  }, []);

  useEffect(() => {
    const prevent = (event: DragEvent) => {
      event.preventDefault();
      setDropVisible(event.type !== "drop" && event.type !== "dragleave");
    };
    const drop = (event: DragEvent) => {
      event.preventDefault();
      setDropVisible(false);
      const file = event.dataTransfer?.files[0];
      if (file) {
        setError("");
        setSrc(file);
      }
    };
    window.addEventListener("dragenter", prevent);
    window.addEventListener("dragover", prevent);
    window.addEventListener("dragleave", prevent);
    window.addEventListener("drop", drop);
    return () => {
      window.removeEventListener("dragenter", prevent);
      window.removeEventListener("dragover", prevent);
      window.removeEventListener("dragleave", prevent);
      window.removeEventListener("drop", drop);
    };
  }, []);

  return (
    <div className="app-shell">
      <header className="toolbar">
        <div className="brand">
          <span className="brand-mark">o</span>
          <div>
            <strong>orim</strong>
            <small>DOCX reader</small>
          </div>
        </div>
        <label className="file-button">
          Open .docx
          <input
            type="file"
            accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            onChange={(event) => {
              const file = event.currentTarget.files?.[0];
              if (file) {
                setError("");
                setSrc(file);
              }
            }}
          />
        </label>
        <div className="segmented" role="group" aria-label="Tracked changes">
          {(["markup", "final", "original"] as const).map((mode) => (
            <button
              className={changes === mode ? "active" : ""}
              data-mode={mode}
              key={mode}
              onClick={() => setChanges(mode)}
              type="button"
            >
              {mode[0].toUpperCase() + mode.slice(1)}
            </button>
          ))}
        </div>
        <label className="comment-toggle">
          <input
            checked={comments}
            onChange={(event) => setComments(event.currentTarget.checked)}
            type="checkbox"
          />
          Comments
        </label>
      </header>
      {error && <div className="empty-state error">{error}</div>}
      {!src && !error && <div className="empty-state">Loading sample document…</div>}
      <DocxViewer
        className="viewer"
        src={src}
        changes={changes}
        comments={comments}
        onError={(reason) => setError(reason instanceof Error ? reason.message : String(reason))}
      />
      <div className={`drop-hint${dropVisible ? " visible" : ""}`}>
        Drop a Word document to open
      </div>
    </div>
  );
}
