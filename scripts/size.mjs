import { readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { gzipSync } from "node:zlib";

const budget = 20 * 1024;
function assets(directory) {
  return readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const path = join(directory, entry.name);
    if (entry.isDirectory()) return assets(path);
    return /\.(?:js|css)$/.test(entry.name) ? [path] : [];
  });
}
const files = assets("dist");
if (!files.some((file) => file.endsWith(".js"))) {
  throw new Error("No packed JavaScript found in dist/. Run vp pack first.");
}
const sizes = files.map((file) => ({
  file,
  raw: readFileSync(file).byteLength,
  gzip: gzipSync(readFileSync(file)).byteLength,
}));
const total = sizes.reduce((sum, item) => sum + item.gzip, 0);
for (const item of sizes) {
  console.log(
    `${item.file}: ${item.raw.toLocaleString()} bytes, ${item.gzip.toLocaleString()} bytes gzip`,
  );
}
console.log(
  `Total packed JS + CSS: ${total.toLocaleString()} bytes gzip (budget ${budget.toLocaleString()})`,
);
if (total > budget) process.exitCode = 1;
