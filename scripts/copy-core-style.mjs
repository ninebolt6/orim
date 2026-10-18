import { copyFileSync } from "node:fs";

copyFileSync(
  new URL("../packages/core/src/docx/render/styles.css", import.meta.url),
  new URL("../packages/core/dist/docx.css", import.meta.url),
);

copyFileSync(
  new URL("../packages/core/src/xlsx/render/styles.css", import.meta.url),
  new URL("../packages/core/dist/xlsx.css", import.meta.url),
);
