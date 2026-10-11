# orim

orim is a small TypeScript DOCX reader for modern browsers. It reads `.docx` files directly in Chrome without a runtime dependency or server conversion.

```ts
import { renderDocx } from "orim";

const rendered = await renderDocx(file, document.querySelector("#document")!, {
  changes: "markup",
  comments: true,
});

rendered.setChanges("final");
rendered.destroy();
```

`parseDocx` can also open a document as a lightweight model. Page dimensions and margins come from section properties; text, tables, images, comments, basic lists, headers/footers, hyperlinks and tracked insertions/deletions are rendered from OOXML parts.

## Limitations

- Layout is page-level rather than line-level: a single block taller than a page can overflow that page. Tables split at row boundaries.
- Floating objects render inline. VML, charts, SmartArt, conditional table-style formatting, paragraph-mark merging and exact Word font metrics are not implemented.
- The ZIP reader supports stored and deflate entries, not ZIP64, encryption or multi-disk archives. Deflated files require `DecompressionStream('deflate-raw')`.
- Page geometry and typography approximate Word; detailed pagination should be checked in Chrome.

## Development

This repository uses Vite+ and pnpm. The project-local CLI works with `pnpm exec vp`.

```sh
pnpm install
pnpm exec vp check
pnpm exec vp test
pnpm exec vp pack
pnpm run size
pnpm exec vp build
pnpm run fixtures
```

The demo loads the committed sample document by default. The fixture generator uses `fflate` as a development-only dependency; the library has no runtime dependencies.
