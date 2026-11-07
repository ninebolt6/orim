import { useEffect, useRef, useState } from "react";
import { DocxViewer, PptxViewer, XlsxViewer } from "@orim/react";
import type { ChangesMode } from "@orim/core";
import { loadSample, parseFile } from "./load";
import type { LoadResult, SampleKind, View } from "./load";

interface AppProps {
  initialLoad: Promise<LoadResult>;
}

export const App = ({ initialLoad }: AppProps) => {
  const [view, setView] = useState<View | null>(null);
  const [changes, setChanges] = useState<ChangesMode>("markup");
  const [comments, setComments] = useState(true);
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(true);
  const [dropVisible, setDropVisible] = useState(false);
  const requestId = useRef(0);

  useEffect(() => {
    let active = true;
    void initialLoad.then((result) => {
      if (!active || requestId.current !== 0) return;
      setView(result.view);
      setError(result.error);
      setLoading(false);
    });
    return () => {
      active = false;
    };
  }, [initialLoad]);

  const open = (file: Blob, extension: string): void => {
    const request = ++requestId.current;
    setError("");
    setLoading(true);
    void parseFile(file, extension).then(
      (parsed) => {
        if (request !== requestId.current) return;
        setView(parsed);
        setLoading(false);
      },
      (reason: unknown) => {
        if (request !== requestId.current) return;
        setError(reason instanceof Error ? reason.message : String(reason));
        setLoading(false);
      },
    );
  };

  const openSample = async (kind: SampleKind): Promise<void> => {
    const request = ++requestId.current;
    setError("");
    setLoading(true);
    try {
      const parsed = await loadSample(kind);
      if (request !== requestId.current) return;
      setView(parsed);
      setLoading(false);
    } catch (reason) {
      if (request !== requestId.current) return;
      setError(reason instanceof Error ? reason.message : String(reason));
      setLoading(false);
    }
  };

  const openFile = (file: File): void => {
    const extension = file.name.includes(".") ? file.name.toLowerCase().split(".").at(-1)! : "";
    open(file, extension);
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
        if (file) openFile(file);
      }}
    >
      <header className="toolbar">
        <div className="brand">
          <span className="brand-mark">o</span>
          <div>
            <strong>orim</strong>
            <small>Document reader</small>
          </div>
        </div>
        <label className="file-button">
          Open file
          <input
            type="file"
            accept=".docx,.xlsx,.pptx"
            onChange={(event) => {
              const file = event.currentTarget.files?.[0];
              event.currentTarget.value = "";
              if (file) openFile(file);
            }}
          />
        </label>
        <button
          className="sample-button"
          data-short-label="Word"
          onClick={() => void openSample("docx")}
          type="button"
        >
          Word sample
        </button>
        <button
          className="sample-button"
          data-short-label="Excel"
          onClick={() => void openSample("xlsx")}
          type="button"
        >
          Excel sample
        </button>
        <button
          className="sample-button"
          data-short-label="PowerPoint"
          onClick={() => void openSample("pptx")}
          type="button"
        >
          PowerPoint sample
        </button>
        {view?.kind === "docx" && (
          <>
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
          </>
        )}
      </header>
      {error && (
        <div className="empty-state error" role="alert">
          {error}
        </div>
      )}
      {loading && (
        <div className="empty-state loading" role="status">
          Loading file…
        </div>
      )}
      {!view && !error && !loading && (
        <div className="empty-state">Open a DOCX, XLSX or PPTX file.</div>
      )}
      {view?.kind === "docx" && (
        <DocxViewer
          className="viewer"
          document={view.document}
          changes={changes}
          comments={comments}
        />
      )}
      {view?.kind === "xlsx" && (
        <XlsxViewer className="viewer viewer-xlsx" workbook={view.workbook} />
      )}
      {view?.kind === "pptx" && (
        <PptxViewer className="viewer viewer-pptx" presentation={view.presentation} />
      )}
      <div className={`drop-hint${dropVisible ? " visible" : ""}`}>
        Drop a Word, Excel or PowerPoint file to open
      </div>
    </div>
  );
};
