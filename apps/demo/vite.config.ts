import { defineConfig } from "vite-plus";

export default defineConfig({
  build: {
    outDir: "dist",
    target: "es2022",
  },
});
