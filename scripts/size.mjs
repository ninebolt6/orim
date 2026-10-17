import { readFileSync, readdirSync } from "node:fs";
import { resolve, join } from "node:path";
import { gzipSync } from "node:zlib";
import { build, minifySync } from "vite-plus";

const assets = (directory) =>
  readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const path = join(directory, entry.name);
    if (entry.isDirectory()) return assets(path);
    return /\.(?:js|css)$/.test(entry.name) ? [path] : [];
  });

const gzip = (content) => gzipSync(content).byteLength;
const coreDist = resolve("packages/core/dist/index.js");
const measureConsumer = async (name, exports, budget) => {
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
      minify: true,
      target: "es2022",
      lib: { entry: entryId, formats: ["es"], fileName: "bundle" },
      rollupOptions: { output: { codeSplitting: false } },
    },
  });
  const outputs = Array.isArray(result) ? result.flatMap((item) => item.output) : result.output;
  const chunks = outputs.filter((item) => item.type === "chunk");
  if (chunks.length !== 1) throw new Error(`Expected one ${name} consumer bundle.`);
  const minified = minifySync(`${name.toLowerCase()}-consumer.js`, chunks[0].code, {
    module: true,
    compress: true,
    mangle: true,
  });
  if (minified.errors.length) throw minified.errors[0];
  const size = gzip(minified.code);
  console.log(
    `${name} consumer bundle: ${size.toLocaleString()} bytes gzip (budget ${budget.toLocaleString()})`,
  );
  return size > budget;
};

const packages = [
  { name: "@orim/core", directory: "packages/core/dist", budget: 32 * 1024 },
  { name: "@orim/react", directory: "packages/react/dist", budget: 2 * 1024 },
];
let failed = false;
for (const pkg of packages) {
  const files = assets(pkg.directory).filter((file) => file.endsWith(".js"));
  if (files.length === 0)
    throw new Error(`No packed JavaScript found in ${pkg.directory}; run pnpm run build first.`);
  const total = files.reduce((sum, file) => sum + gzip(readFileSync(file)), 0);
  console.log(
    `${pkg.name} dist JS: ${total.toLocaleString()} bytes gzip (budget ${pkg.budget.toLocaleString()})`,
  );
  if (total > pkg.budget) failed = true;
}
const docxFailed = await measureConsumer(
  "DOCX",
  ["parseDocx", "renderDocument", "renderDocx"],
  20 * 1024,
);
const xlsxFailed = await measureConsumer(
  "XLSX",
  ["parseXlsx", "renderWorkbook", "renderXlsx"],
  14 * 1024,
);
failed = failed || docxFailed || xlsxFailed;
for (const [name, path, budget] of [
  ["docx.css", "packages/core/dist/docx.css", 4 * 1024],
  ["xlsx.css", "packages/core/dist/xlsx.css", 2 * 1024],
]) {
  const size = gzip(readFileSync(path));
  console.log(`${name}: ${size.toLocaleString()} bytes gzip (budget ${budget.toLocaleString()})`);
  if (size > budget) failed = true;
}
if (failed) process.exitCode = 1;
