import { useRef, useState } from "react";
import { DocxViewer } from "@orim/react";
import { parseDocx } from "@orim/core";
import type { ChangesMode, DocxDocument } from "@orim/core";

interface AppProps {
  initialDocument: DocxDocument | null;
  initialError: string;
}

export const App = ({ initialDocument, initialError }: AppProps) => {
  const [document, setDocument] = useState(initialDocument);
  const [changes, setChanges] = useState<ChangesMode>("markup");
  const [comments, setComments] = useState(true);
  const [error, setError] = useState(initialError);
  const [loading, setLoading] = useState(false);
  const [dropVisible, setDropVisible] = useState(false);
  const requestId = useRef(0);

  const open = (file: File) => {
    const request = ++requestId.current;
    setError("");
    setLoading(true);
    void parseDocx(file).then(
      (parsed) => {
        if (request !== requestId.current) return;
        setDocument(parsed);
        setLoading(false);
      },
      (reason: unknown) => {
        if (request !== requestId.current) return;
        setError(reason instanceof Error ? reason.message : String(reason));
        setLoading(false);
      },
    );
  };

  return (
    <div
      className="app-shell"
      onDragEnter={(event) => {
        event.preventDefault();
        setDropVisible(true);
      }}
      onDragOver={(event) => {
        event.preventDefault();
        event.dataTransfer.dropEffect = "copy";
      }}
      onDragLeave={(event) => {
        event.preventDefault();
        if (
          !(
            event.relatedTarget instanceof Node && event.currentTarget.contains(event.relatedTarget)
          )
        )
          setDropVisible(false);
      }}
      onDrop={(event) => {
        event.preventDefault();
        setDropVisible(false);
        const file = event.dataTransfer.files[0];
        if (file) open(file);
      }}
    >
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
              event.currentTarget.value = "";
              if (file) open(file);
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
      {error && (
        <div className="empty-state error" role="alert">
          {error}
        </div>
      )}
      {loading && (
        <div className="empty-state loading" role="status">
          Loading document…
        </div>
      )}
      {!document && !error && !loading && (
        <div className="empty-state">Open a .docx document to begin.</div>
      )}
      <DocxViewer className="viewer" document={document} changes={changes} comments={comments} />
      <div className={`drop-hint${dropVisible ? " visible" : ""}`}>
        Drop a Word document to open
      </div>
    </div>
  );
};
