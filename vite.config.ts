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
        files: [
          "packages/core/src/docx/*.ts",
          "packages/core/src/docx/parse/**",
          "packages/core/src/xlsx/*.ts",
          "packages/core/src/xlsx/parse/**",
          "packages/core/src/pptx/*.ts",
          "packages/core/src/pptx/parse/**",
          "packages/core/src/drawingml/**",
          "packages/core/src/opc/**",
          "packages/core/src/common/xml.ts",
        ],
        rules: {
          "no-param-reassign": ["error", { props: true }],
        },
      },
      {
        files: ["packages/core/src/docx/**"],
        rules: {
          "no-restricted-imports": ["error", { patterns: ["**/xlsx/**", "**/pptx/**"] }],
        },
      },
      {
        files: ["packages/core/src/xlsx/**"],
        rules: {
          "no-restricted-imports": ["error", { patterns: ["**/docx/**", "**/pptx/**"] }],
        },
      },
      {
        files: ["packages/core/src/pptx/**"],
        rules: {
          "no-restricted-imports": ["error", { patterns: ["**/docx/**", "**/xlsx/**"] }],
        },
      },
      {
        files: ["packages/core/src/drawingml/**"],
        rules: {
          "no-restricted-imports": [
            "error",
            { patterns: ["**/docx/**", "**/xlsx/**", "**/pptx/**"] },
          ],
        },
      },
      {
        files: ["packages/core/src/common/**", "packages/core/src/opc/**"],
        rules: {
          "no-restricted-imports": [
            "error",
            { patterns: ["**/docx/**", "**/xlsx/**", "**/pptx/**", "**/drawingml/**"] },
          ],
        },
      },
    ],
  },
  test: {
    coverage: {
      provider: "v8",
      include: ["packages/*/src/**/*.{ts,tsx}"],
      exclude: [
        "**/*.test.{ts,tsx}",
        "**/*.integration.test.{ts,tsx}",
        "**/*.browser.test.{ts,tsx}",
        "**/*.d.ts",
        "**/index.ts",
      ],
      reporter: ["text", "json"],
    },
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
