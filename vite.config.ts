import { defineConfig } from "vite-plus";
import { playwright } from "vite-plus/test/browser-playwright";

export default defineConfig({
  lint: {
    rules: {
      "func-style": ["error", "expression"],
    },
    overrides: [
      {
        files: ["packages/react/**", "apps/demo/**"],
        plugins: ["react", "react-hooks"],
        rules: {
          "react-hooks/rules-of-hooks": "error",
          "react-hooks/exhaustive-deps": "error",
        },
      },
      {
        files: ["packages/core/src/docx/**", "packages/core/src/opc/**"],
        rules: {
          "no-param-reassign": ["error", { props: true }],
        },
      },
    ],
  },
  test: {
    projects: [
      {
        test: {
          name: "unit",
          environment: "happy-dom",
          setupFiles: ["./packages/core/test/namespace-dom.ts"],
          include: ["packages/**/*.test.{ts,tsx}", "apps/**/*.test.{ts,tsx}"],
          exclude: ["**/node_modules/**", "**/dist/**", "**/*.browser.test.{ts,tsx}"],
        },
      },
      {
        test: {
          name: "browser",
          include: ["packages/**/*.browser.test.{ts,tsx}", "apps/**/*.browser.test.{ts,tsx}"],
          browser: {
            enabled: true,
            provider: playwright(),
            headless: true,
            instances: [{ browser: "chromium" }],
          },
        },
      },
    ],
  },
  staged: {
    "*": "vp check --fix",
  },
});
