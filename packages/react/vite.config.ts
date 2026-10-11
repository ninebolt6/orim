import { defineConfig } from "vite-plus";

export default defineConfig({
  pack: {
    entry: ["src/index.ts"],
    format: "esm",
    dts: true,
    minify: true,
    outDir: "dist",
    platform: "browser",
    target: "es2022",
    external: ["react", "react-dom", "@orim/core"],
    jsx: "react-jsx",
  },
});
