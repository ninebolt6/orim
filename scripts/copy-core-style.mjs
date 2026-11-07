import { copyFileSync } from "node:fs";

for (const format of ["docx", "xlsx", "pptx"]) {
  copyFileSync(
    new URL(`../packages/core/src/${format}/render/styles.css`, import.meta.url),
    new URL(`../packages/core/dist/${format}.css`, import.meta.url),
  );
}
