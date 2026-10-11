import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { parseDocx } from "@orim/core";
import type { DocxDocument } from "@orim/core";
import "@orim/core/style.css";
import { App } from "./App";
import "./style.css";
import sampleUrl from "../../../fixtures/sample.docx?url";

let initialDocument: DocxDocument | null = null;
let initialError = "";

try {
  const response = await fetch(sampleUrl);
  if (!response.ok) throw new Error(`Unable to load sample: ${response.status}`);
  initialDocument = await parseDocx(await response.blob());
} catch (reason) {
  initialError = reason instanceof Error ? reason.message : String(reason);
}

createRoot(document.getElementById("app")!).render(
  <StrictMode>
    <App initialDocument={initialDocument} initialError={initialError} />
  </StrictMode>,
);
