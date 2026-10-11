import { describe, expect, it } from "vite-plus/test";
import { fontStack } from "./text-style";

describe("font fallbacks", () => {
  it("adds Japanese Mincho and Gothic system families", () => {
    expect(fontStack("ＭＳ 明朝")).toContain("Noto Serif CJK JP");
    expect(fontStack("Meiryo")).toContain("Noto Sans CJK JP");
    expect(fontStack("Times New Roman")).toContain("serif");
    expect(fontStack("Georgia")).toMatch(/Georgia.*serif$/);
    expect(fontStack("Arial")).toMatch(/Arial.*sans-serif$/);
  });
});
