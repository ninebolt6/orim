import { readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { gzipSync } from "node:zlib";

function assets(directory) {
  return readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const path = join(directory, entry.name);
    if (entry.isDirectory()) return assets(path);
    return /\.(?:js|css)$/.test(entry.name) ? [path] : [];
  });
}

const packages = [
  { name: "@orim/core", directory: "packages/core/dist", budget: 20 * 1024 },
  { name: "@orim/react", directory: "packages/react/dist", budget: 2 * 1024 },
];
let failed = false;
for (const pkg of packages) {
  const files = assets(pkg.directory).filter((file) => file.endsWith(".js"));
  if (!files.some((file) => file.endsWith(".js")))
    throw new Error(`No packed JavaScript found in ${pkg.directory}; run pnpm run build first.`);
  const total = files.reduce((sum, file) => sum + gzipSync(readFileSync(file)).byteLength, 0);
  console.log(
    `${pkg.name} JS: ${total.toLocaleString()} bytes gzip (budget ${pkg.budget.toLocaleString()})`,
  );
  if (total > pkg.budget) failed = true;
}
const cssPath = "packages/core/dist/style.css";
try {
  const cssSize = gzipSync(readFileSync(cssPath)).byteLength;
  console.log(
    `@orim/core CSS: ${cssSize.toLocaleString()} bytes gzip (budget ${(
      4 * 1024
    ).toLocaleString()})`,
  );
  if (cssSize > 4 * 1024) failed = true;
} catch {
  throw new Error(`No published CSS found at ${cssPath}; run pnpm run build first.`);
}
if (failed) process.exitCode = 1;
