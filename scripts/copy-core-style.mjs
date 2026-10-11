import { copyFileSync } from "node:fs";

copyFileSync(
  new URL("../packages/core/src/render/styles.css", import.meta.url),
  new URL("../packages/core/dist/style.css", import.meta.url),
);
