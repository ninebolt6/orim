import { describe, expect, it } from "vite-plus/test";
import { xmlDocument } from "../../test/xml";
import { indexedColors, parseColor, themeColors } from "./colors";

describe("XLSX colors", () => {
  it("reads theme, RGB, and indexed colors", () => {
    const theme = themeColors(
      xmlDocument(
        '<a:theme xmlns:a="urn:a"><a:themeElements><a:clrScheme><a:lt1><a:sysClr lastClr="F1F2F3"/></a:lt1><a:dk1><a:srgbClr val="010203"/></a:dk1></a:clrScheme></a:theme>',
      ),
    );
    const styles = xmlDocument(
      '<styleSheet><colors><indexedColors><rgbColor rgb="FFAABBCC"/></indexedColors></colors></styleSheet>',
    );
    const palette = indexedColors(styles);
    expect(theme.slice(0, 2)).toEqual(["#f1f2f3", "#010203"]);
    expect(parseColor(xmlDocument('<color rgb="FF123456"/>').documentElement, theme, palette)).toBe(
      "#123456",
    );
    expect(parseColor(xmlDocument('<color theme="1"/>').documentElement, theme, palette)).toBe(
      "#010203",
    );
    expect(parseColor(xmlDocument('<color indexed="0"/>').documentElement, theme, palette)).toBe(
      "#aabbcc",
    );
  });

  it("applies HLS tints and system colors", () => {
    expect(parseColor(xmlDocument('<color rgb="FF000000" tint="0.5"/>').documentElement)).toBe(
      "#808080",
    );
    expect(parseColor(xmlDocument('<color rgb="FFFFFFFF" tint="-0.5"/>').documentElement)).toBe(
      "#808080",
    );
    expect(parseColor(xmlDocument('<color indexed="64"/>').documentElement)).toBe("#000000");
    expect(parseColor(xmlDocument('<color indexed="65"/>').documentElement)).toBe("#ffffff");
    expect(
      parseColor(xmlDocument('<color auto="1" rgb="FF123456"/>').documentElement),
    ).toBeUndefined();
  });
});
