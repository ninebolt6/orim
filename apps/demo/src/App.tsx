import { useRef, useState } from "react";
import { DocxViewer, XlsxViewer } from "@orim/react";
import { parseDocx, parseXlsx } from "@orim/core";
import type { ChangesMode, DocxDocument, XlsxWorkbook } from "@orim/core";
import sampleDocxUrl from "../../../fixtures/sample.docx?url";
import sampleXlsxUrl from "../../../fixtures/sample.xlsx?url";

type View =
  | { readonly kind: "docx"; readonly document: DocxDocument }
  | { readonly kind: "xlsx"; readonly workbook: XlsxWorkbook };

interface AppProps {
  initialView: View | null;
  initialError: string;
}

const parseFile = async (file: Blob, extension: string): Promise<View> =>
  extension === "xlsx"
    ? { kind: "xlsx", workbook: await parseXlsx(file) }
    : { kind: "docx", document: await parseDocx(file) };

export const App = ({ initialView, initialError }: AppProps) => {
  const [view, setView] = useState<View | null>(initialView);
  const [changes, setChanges] = useState<ChangesMode>("markup");
  const [comments, setComments] = useState(true);
  const [error, setError] = useState(initialError);
  const [loading, setLoading] = useState(false);
  const [dropVisible, setDropVisible] = useState(false);
  const requestId = useRef(0);

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

  const openSample = async (kind: "docx" | "xlsx"): Promise<void> => {
    const url = kind === "docx" ? sampleDocxUrl : sampleXlsxUrl;
    const request = ++requestId.current;
    setError("");
    setLoading(true);
    try {
      const response = await fetch(url);
      if (!response.ok) throw new Error(`Unable to load sample: ${response.status}`);
      const parsed = await parseFile(await response.blob(), kind);
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
    const extension = file.name.toLowerCase().split(".").at(-1) === "xlsx" ? "xlsx" : "docx";
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
            accept=".docx,.xlsx,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            onChange={(event) => {
              const file = event.currentTarget.files?.[0];
              event.currentTarget.value = "";
              if (file) openFile(file);
            }}
          />
        </label>
        <button className="sample-button" onClick={() => void openSample("docx")} type="button">
          Word sample
        </button>
        <button className="sample-button" onClick={() => void openSample("xlsx")} type="button">
          Excel sample
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
      {!view && !error && !loading && <div className="empty-state">Open a DOCX or XLSX file.</div>}
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
      <div className={`drop-hint${dropVisible ? " visible" : ""}`}>
        Drop a Word or Excel file to open
      </div>
    </div>
  );
};
