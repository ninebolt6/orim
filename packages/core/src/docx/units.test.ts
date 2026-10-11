import { describe, expect, it } from "vite-plus/test";
import { cssWidth, twips } from "./units";
import { W } from "../xml";

function width(value: string, type: string): Element {
  return new DOMParser().parseFromString(
    `<root xmlns:w="${W}"><w:tblW w:w="${value}" w:type="${type}"/></root>`,
    "application/xml",
  ).documentElement!.firstElementChild!;
}

describe("DOCX units", () => {
  it("converts twips and resolves width units", () => {
    expect(twips("360")).toBe(18);
    expect(twips(undefined, 7)).toBe(7);
    expect(cssWidth(width("1440", "dxa"))).toBe("72pt");
    expect(cssWidth(width("2500", "pct"))).toBe("50%");
    expect(cssWidth(width("50%", "pct"))).toBe("50%");
    expect(cssWidth(width("0", "auto"))).toBeUndefined();
  });
});
