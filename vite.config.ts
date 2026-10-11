import { defineConfig } from "vite-plus";

export default defineConfig({
  build: {
    outDir: "demo-dist",
  },
  pack: {
    entry: ["src/index.ts"],
    format: "esm",
    dts: true,
    minify: true,
    outDir: "dist",
    platform: "browser",
    target: "es2022",
  },
  test: {
    environment: "happy-dom",
    include: ["src/**/*.test.ts"],
  },
  staged: {
    "*": "vp check --fix",
  },
});
