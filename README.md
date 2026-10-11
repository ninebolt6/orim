# orim

orim is a small TypeScript DOCX reader for modern browsers. It reads `.docx` files directly in the browser without server conversion or runtime dependencies.

## Packages

- `@orim/core` parses OOXML and renders paginated documents into a DOM container.
- `@orim/react` provides the `DocxViewer` React component and treats React as a peer dependency.

The workspace demo in `apps/demo` shows the renderer with the generated sample document.

## Vanilla JavaScript

```ts
import { parseDocx, renderDocx } from "@orim/core";

const rendered = await renderDocx(file, document.querySelector("#document")!, {
  changes: "markup",
  comments: true,
});

rendered.update({ changes: "final" });
rendered.destroy();

const model = await parseDocx(file);
```

## React

```tsx
import { DocxViewer } from "@orim/react";

<DocxViewer src={file} changes="markup" comments />;
```

Pass an `ArrayBuffer`, `Uint8Array`, `Blob`, parsed `DocxDocument`, or `null` to `src`. Page dimensions and margins come from section properties; text, tables, images, comments, basic lists, headers/footers, hyperlinks and tracked changes are rendered from OOXML parts.

## Limitations

- Layout is page-level rather than line-level: a single block taller than a page can overflow that page. Tables split at row boundaries.
- Floating objects render inline. VML, charts, SmartArt, conditional table-style formatting, paragraph-mark merging and exact Word font metrics are not implemented.
- The ZIP reader supports stored and deflate entries, not ZIP64, encryption or multi-disk archives. Deflated files require `DecompressionStream('deflate-raw')`.
- Deflated parts are limited to 64 MiB uncompressed and must not exceed their declared size.
- Page geometry and typography approximate Word; detailed pagination should be checked in Chrome.

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
