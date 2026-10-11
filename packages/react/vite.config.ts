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
    inputOptions: {
      external: ["react", "react-dom", "react/jsx-runtime", "@orim/core"],
    },
    deps: {
      neverBundle: true,
    },
    jsx: "react-jsx",
  },
});
