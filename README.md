# orim

orim is a small TypeScript DOCX and XLSX viewer for modern browsers. It reads `.docx` and `.xlsx` files directly in the browser without server conversion or runtime dependencies.

## Packages

- `@orim/core` parses OOXML and renders DOCX documents and XLSX worksheets into a DOM container.
- `@orim/react` provides the `DocxViewer` and `XlsxViewer` components and treats React as a peer dependency.

Import `@orim/core/docx.css` and/or `@orim/core/xlsx.css` once to apply the corresponding renderer styles.

The workspace demo in `apps/demo` shows both formats with generated Word and spreadsheet samples.

## DOCX (vanilla)

Add the controls below, then run the TypeScript in your browser entry module.

```html
<input id="file" type="file" accept=".docx" aria-label="Open DOCX" />
<p id="error" role="alert"></p>
<div id="viewer"></div>
```

```ts
import { renderDocx } from "@orim/core";
import type { RenderedDocument } from "@orim/core";
import "@orim/core/docx.css";

const input = document.querySelector<HTMLInputElement>("#file")!;
const container = document.querySelector<HTMLElement>("#viewer")!;
const error = document.querySelector<HTMLElement>("#error")!;
const controller = new AbortController();
let rendered: RenderedDocument | undefined;

input.addEventListener("change", async () => {
  const file = input.files?.[0];
  input.value = "";
  if (!file) return;
  input.disabled = true;
  error.textContent = "";
  try {
    rendered = await renderDocx(file, container, {
      changes: "markup",
      comments: true,
      signal: controller.signal,
    });
  } catch (reason) {
    if (!controller.signal.aborted)
      error.textContent = reason instanceof Error ? reason.message : String(reason);
  } finally {
    input.disabled = false;
  }
});

window.addEventListener("pagehide", () => {
  controller.abort();
  rendered?.destroy();
});
```

The handle stays alive while viewing; use `rendered.update({ changes: "final" })` to change revision display. Call `destroy()` when your view is torn down (the example uses `pagehide`); aborting also cancels a pending render.

## XLSX (vanilla)

This is a separate page/example with its own controls.

```html
<input id="file" type="file" accept=".xlsx" aria-label="Open XLSX" />
<p id="error" role="alert"></p>
<div id="viewer"></div>
```

```ts
import { renderXlsx } from "@orim/core";
import type { RenderedWorkbook } from "@orim/core";
import "@orim/core/xlsx.css";

const input = document.querySelector<HTMLInputElement>("#file")!;
const container = document.querySelector<HTMLElement>("#viewer")!;
const error = document.querySelector<HTMLElement>("#error")!;
const controller = new AbortController();
let rendered: RenderedWorkbook | undefined;

input.addEventListener("change", async () => {
  const file = input.files?.[0];
  input.value = "";
  if (!file) return;
  input.disabled = true;
  error.textContent = "";
  try {
    rendered = await renderXlsx(file, container, { signal: controller.signal });
  } catch (reason) {
    if (!controller.signal.aborted)
      error.textContent = reason instanceof Error ? reason.message : String(reason);
  } finally {
    input.disabled = false;
  }
});

window.addEventListener("pagehide", () => {
  controller.abort();
  rendered?.destroy();
});
```

Worksheet tabs switch sheets. With a tab focused, ArrowLeft/ArrowRight wrap between visible sheets; Home/End select the first/last. This also works in React. To select a sheet programmatically, use `rendered.update({ sheet: index })` with a zero-based, visible worksheet index.

## React

Render this component from your React application. Parsing happens in the file-change handler; the viewers receive parsed models and clean up their rendering on unmount.

```tsx
import { useRef, useState } from "react";
import type { ChangeEvent } from "react";
import { parseDocx, parseXlsx } from "@orim/core";
import type { DocxDocument, XlsxWorkbook } from "@orim/core";
import { DocxViewer, XlsxViewer } from "@orim/react";
import "@orim/core/docx.css";
import "@orim/core/xlsx.css";

type View = { kind: "docx"; document: DocxDocument } | { kind: "xlsx"; workbook: XlsxWorkbook };

export const OfficeViewer = () => {
  const [view, setView] = useState<View | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const requestId = useRef(0);

  const openFile = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.currentTarget.files?.[0];
    event.currentTarget.value = "";
    if (!file) return;
    const request = ++requestId.current;
    setLoading(true);
    setError("");
    try {
      const extension = file.name.toLowerCase().match(/\.([^.]+)$/)?.[1];
      let next: View;
      if (extension === "docx") next = { kind: "docx", document: await parseDocx(file) };
      else if (extension === "xlsx") next = { kind: "xlsx", workbook: await parseXlsx(file) };
      else throw new Error("Open a .docx or .xlsx file.");
      if (request === requestId.current) setView(next);
    } catch (reason) {
      if (request === requestId.current)
        setError(reason instanceof Error ? reason.message : String(reason));
    } finally {
      if (request === requestId.current) setLoading(false);
    }
  };

  return (
    <>
      <input type="file" accept=".docx,.xlsx" aria-label="Open Office file" onChange={openFile} />
      {loading && <p role="status">Loading file…</p>}
      {error && <p role="alert">{error}</p>}
      {view?.kind === "docx" && <DocxViewer document={view.document} changes="markup" comments />}
      {view?.kind === "xlsx" && <XlsxViewer workbook={view.workbook} />}
    </>
  );
};
```

## What is rendered

DOCX page dimensions and margins come from section properties; text, tables, images, comments, basic lists, headers/footers, hyperlinks and tracked changes are rendered from OOXML parts. XLSX rendering includes worksheets, sparse cells, styles, merges, frozen panes and safe external hyperlinks.

## Supported files / errors

Only `.docx` and `.xlsx` are supported. Legacy `.doc`/`.xls` and password-protected Office containers throw a clear error asking you to save as `.docx`/`.xlsx` without a password. Catch parsing/rendering errors and show their messages to the user, as above; the file picker's `accept` filter is not validation. The React example keeps the previous document visible on error and ignores results from older requests.

## Limitations

- Layout is page-level rather than line-level: a single block taller than a page can overflow that page. Tables split at row boundaries.
- Floating objects render inline. VML, charts, SmartArt, conditional table-style formatting, paragraph-mark merging and exact Word font metrics are not implemented.
- The ZIP reader supports stored and deflate entries, not ZIP64, encryption or multi-disk archives. Deflated files require `DecompressionStream('deflate-raw')`.
- Deflated parts are limited to 64 MiB uncompressed and must not exceed their declared size.
- Page geometry and typography approximate Word; detailed pagination should be checked in Chrome.
- XLSX display is a worksheet grid rather than a full Excel calculation engine. Formula cells show cached values; unsupported or invalid number formats fall back to General. Rendering is capped at 5,000 rows and 200 columns per worksheet.

## Development

This repository uses Vite+ and pnpm. The project-local CLI works with `pnpm exec vp`.

```sh
pnpm install
pnpm exec vp check
pnpm exec vp test
pnpm run build
pnpm run size
pnpm --filter @orim/demo dev
pnpm run fixtures
```

The demo loads the committed sample document by default. The fixture generator uses `fflate` as a development-only dependency; the core package has no runtime dependencies.
