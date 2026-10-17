import { describe, expect, it } from "vite-plus/test";
import { columnName, parseCellRef, parseRange } from "./ref";

describe("Excel references", () => {
  it("converts zero-based column indexes to names", () => {
    expect(columnName(0)).toBe("A");
    expect(columnName(25)).toBe("Z");
    expect(columnName(26)).toBe("AA");
    expect(columnName(16383)).toBe("XFD");
  });

  it("parses bounded cell references", () => {
    expect(parseCellRef("AB12")).toEqual({ row: 11, column: 27 });
    expect(parseCellRef("$AB$12")).toEqual({ row: 11, column: 27 });
    expect(parseCellRef("XFE1")).toBeUndefined();
    expect(parseCellRef("A1048577")).toBeUndefined();
    expect(parseCellRef("A0")).toBeUndefined();
    expect(parseCellRef("A1:B2")).toBeUndefined();
  });

  it("normalizes single-cell and rectangular ranges", () => {
    expect(parseRange("A1:B2")).toEqual({ top: 0, left: 0, bottom: 1, right: 1 });
    expect(parseRange("C3")).toEqual({ top: 2, left: 2, bottom: 2, right: 2 });
    expect(parseRange("B2:A1")).toEqual({ top: 0, left: 0, bottom: 1, right: 1 });
    expect(parseRange("A1:B2:C3")).toBeUndefined();
    expect(parseRange("A1:XFE1")).toBeUndefined();
  });
});
