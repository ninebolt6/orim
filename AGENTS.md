# Using Vite+

This project uses Vite+, a unified toolchain for the web. Vite+ is available locally as `pnpm exec vp`; use it for development, checking, testing, building and packing.

- Run `pnpm exec vp check` and `pnpm exec vp test` before handing off changes.
- Use `pnpm exec vp build` for the demo and `pnpm exec vp pack` for the library.
- Keep the single root `vite.config.ts` as the source of Vite+, test, pack and staged-file configuration.

# orim notes

- Keep the published library free of runtime dependencies. `fflate` is only for fixture generation and tests.
- Keep OOXML parsing and browser rendering separate; do not add spreadsheet or presentation readers.
- Create document text only with DOM text nodes; sanitize all document-provided links and colors.
- Run the pre-commit checks through `pnpm exec vp staged`.
