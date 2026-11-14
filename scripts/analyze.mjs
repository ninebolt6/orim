import { resolve } from "node:path";
import { gzipSync } from "node:zlib";
import { build } from "vite-plus";

const coreDist = resolve("packages/core/src/index.ts");
const measure = async (name, exports) => {
  const entryId = resolve(`virtual:orim-${name}`);
  const virtualEntryId = `\0${entryId}`;
  const source = `export { ${exports.join(", ")} } from ${JSON.stringify(coreDist)};`;
  const result = await build({
    configFile: false,
    logLevel: "silent",
    plugins: [
      {
        name: `orim-${name}-size-entry`,
        resolveId: (id) => (id === entryId ? virtualEntryId : null),
        load: (id) => (id === virtualEntryId ? source : null),
      },
    ],
    build: {
      write: false,
      minify: false,
      target: "es2022",
      lib: { entry: entryId, formats: ["es"], fileName: "bundle" },
      rollupOptions: { output: { codeSplitting: false } },
    },
  });
  const outputs = Array.isArray(result) ? result.flatMap((item) => item.output) : result.output;
  const chunk = outputs.find((item) => item.type === "chunk");
  const rows = Object.entries(chunk.modules)
    .map(([id, info]) => [id.replace(/^.*\/src\//, ""), info.renderedLength])
    .sort((a, b) => b[1] - a[1]);
  console.log(
    `\n=== ${name} (${chunk.code.length} bytes minified-src, gzip ${gzipSync(chunk.code).byteLength}) ===`,
  );
  for (const [id, len] of rows) console.log(`${String(len).padStart(7)}  ${id}`);
};

await measure("DOCX", ["parseDocx", "renderDocument", "renderDocx"]);
await measure("XLSX", ["parseXlsx", "renderWorkbook", "renderXlsx"]);
await measure("PPTX", ["parsePptx", "renderPresentation", "renderPptx"]);
