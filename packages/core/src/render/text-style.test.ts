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

  it("drops unsafe family names instead of serializing CSS", () => {
    expect(fontStack('Good, "bad')).toContain("Good");
    expect(fontStack('bad";color:red')).toBeUndefined();
    for (const value of [
      "url(x)",
      "image(x)",
      "expression(x)",
      "bad\\family",
      "@import",
      "bad;display:none",
      "<tag>",
      "bad/*comment*/",
    ])
      expect(fontStack(value)).toBeUndefined();
    expect(fontStack("Fancy Name")).toBe('"Fancy Name", sans-serif');
  });
});
