import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import "@orim/core/docx.css";
import "@orim/core/xlsx.css";
import "@orim/core/pptx.css";
import { App } from "./App";
import "./style.css";
import { loadSample } from "./load";
import type { LoadResult } from "./load";

const initialLoad: Promise<LoadResult> = loadSample("docx").then(
  (view) => ({ view, error: "" }),
  (reason: unknown) => ({
    view: null,
    error: reason instanceof Error ? reason.message : String(reason),
  }),
);

createRoot(document.getElementById("app")!).render(
  <StrictMode>
    <App initialLoad={initialLoad} />
  </StrictMode>,
);
