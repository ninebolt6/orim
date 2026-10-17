# orim

orim is a small TypeScript DOCX and XLSX viewer for modern browsers. It reads `.docx` and `.xlsx` files directly in the browser without server conversion or runtime dependencies.

## Packages

- `@orim/core` parses OOXML and renders DOCX documents and XLSX worksheets into a DOM container.
- `@orim/react` provides the `DocxViewer` and `XlsxViewer` components and treats React as a peer dependency.

Import `@orim/core/docx.css` and/or `@orim/core/xlsx.css` once to apply the corresponding renderer styles.

The workspace demo in `apps/demo` shows both formats with generated Word and spreadsheet samples.

## Vanilla JavaScript

```ts
import { parseDocx, renderDocx } from "@orim/core";
import { parseXlsx, renderXlsx, renderWorkbook } from "@orim/core";
import "@orim/core/docx.css";
import "@orim/core/xlsx.css";

const rendered = await renderDocx(file, document.querySelector("#document")!, {
  changes: "markup",
  comments: true,
});

rendered.update({ changes: "final" });
rendered.destroy();

const model = await parseDocx(file);

const workbook = await parseXlsx(file);
const spreadsheet = renderWorkbook(workbook, document.querySelector("#spreadsheet")!);
spreadsheet.update({ sheet: 1 });
spreadsheet.destroy();

const renderedWorkbook = await renderXlsx(file, document.querySelector("#spreadsheet")!);
renderedWorkbook.destroy();
```

## React

```tsx
import { parseDocx } from "@orim/core";
import { parseXlsx } from "@orim/core";
import { DocxViewer, XlsxViewer } from "@orim/react";
import "@orim/core/docx.css";
import "@orim/core/xlsx.css";

const model = await parseDocx(file);

<DocxViewer document={model} changes="markup" comments />;

const workbook = await parseXlsx(file);

<XlsxViewer workbook={workbook} />;
```

Parse DOCX files with `parseDocx` and pass the resulting model to the `document` prop. Page dimensions and margins come from section properties; text, tables, images, comments, basic lists, headers/footers, hyperlinks and tracked changes are rendered from OOXML parts. Parse XLSX files with `parseXlsx` and pass the workbook to `XlsxViewer`; worksheets, sparse cells, styles, merges, frozen panes and safe external hyperlinks are rendered from OOXML parts.

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
