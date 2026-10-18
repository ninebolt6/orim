import { describe, expect, it } from "vite-plus/test";
import { applyTextStyle, changeColor, fontStack } from "./text-style";

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

  it("applies inline font, decoration, color, and visibility styles", () => {
    const element = document.createElement("span");
    applyTextStyle(element, {
      bold: false,
      italic: true,
      underline: "wave",
      strike: true,
      doubleStrike: true,
      color: "#123456",
      highlight: "#abcdef",
      fontFamily: "Arial, serif",
      fontSize: 12,
      verticalAlign: "superscript",
      caps: true,
      smallCaps: true,
      hidden: true,
    });

    expect(element.style.fontWeight).toBe("400");
    expect(element.style.fontStyle).toBe("italic");
    expect(element.style.textDecoration).toContain("double");
    expect(element.style.color).toBe("#123456");
    expect(element.style.backgroundColor).toBe("#abcdef");
    expect(element.style.fontFamily).toContain("Arial");
    expect(element.style.verticalAlign).toBe("super");
    expect(element.style.fontSize).toBe("7.8pt");
    expect(element.style.textTransform).toBe("uppercase");
    expect(element.style.fontVariant).toBe("small-caps");
    expect(element.style.display).toBe("none");
  });

  it("normalizes underline styles and assigns repeatable change colors", () => {
    const underline = document.createElement("span");
    applyTextStyle(underline, { underline: "dotted" });
    expect(underline.style.textDecorationStyle).toBe("dotted");

    const colors = new Map<string, string>();
    const change = { kind: "ins" as const, author: "Aiko", date: "" };
    expect(changeColor(change, colors)).toBe("#137333");
    expect(changeColor(change, colors)).toBe("#137333");
    expect(changeColor({ ...change, author: "Morgan" }, colors)).toBe("#c5221f");
  });
});
