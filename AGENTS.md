# Using Vite+

This project uses Vite+, a unified toolchain for the web. Vite+ is available locally as `pnpm exec vp`; use it for development, checking, testing, building and packing.

- Run `pnpm exec vp check` and `pnpm exec vp test` before handing off changes.
- Use `pnpm exec vp run -r build` for workspace packages and the demo. Package-local `vite.config.ts` files define pack settings.
- Keep shared test, lint, and staged-file configuration in the root `vite.config.ts`.

# orim notes

- Keep the published library free of runtime dependencies. `fflate` is only for fixture generation and tests.
- `docx/`, `xlsx/` and `pptx/` are independent format boundaries and may share only `common/`, `opc/` and `drawingml/` (theme, color, fill/line, EMU and blip parsing shared by all three). `drawingml/` may depend only on `common/` and `opc/`. Import `@orim/core/docx.css`, `@orim/core/xlsx.css` and `@orim/core/pptx.css` for the respective renderers.
- Create document text only with DOM text nodes; sanitize all document-provided links and colors.
- Parsed models are immutable and built with expressions/pure functions. Mutate only DOM construction/pagination, render resources/leases/color maps, memo caches, closure-contained sequential parse state, and local non-escaping ZIP/XML/path accumulators.
- Run the pre-commit checks through `pnpm exec vp staged`.
