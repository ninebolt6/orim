import { describe, expect, it } from "vite-plus/test";
import { xmlDocument } from "../../test/xml";
import { parseColor } from "./color";

const scheme = (name: string): string | undefined => ({ accent1: "#4f81bd", dk1: "#000000" })[name];
const color = (xml: string) =>
  parseColor(
    xmlDocument(
      `<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${xml}</a:solidFill>`,
    ).documentElement,
    scheme,
  );

describe("DrawingML colors", () => {
  it("resolves base color kinds", () => {
    expect(color('<a:srgbClr val="ABCDEF"/>')).toBe("#abcdef");
    expect(color('<a:schemeClr val="accent1"/>')).toBe("#4f81bd");
    expect(color('<a:sysClr val="windowText" lastClr="112233"/>')).toBe("#112233");
    expect(color('<a:prstClr val="black"/>')).toBe("#000000");
    expect(color('<a:schemeClr val="missing"/>')).toBeUndefined();
    expect(color("")).toBeUndefined();
  });

  it("applies tint and shade in linear RGB and HSL modifiers", () => {
    expect(color('<a:schemeClr val="dk1"><a:tint val="75000"/></a:schemeClr>')).toBe("#898989");
    expect(color('<a:srgbClr val="FFFFFF"><a:shade val="50000"/></a:srgbClr>')).toBe("#bcbcbc");
    expect(color('<a:srgbClr val="4F81BD"><a:lumMod val="75000"/></a:srgbClr>')).toBe("#376092");
    expect(
      color('<a:srgbClr val="4F81BD"><a:lumMod val="40000"/><a:lumOff val="60000"/></a:srgbClr>'),
    ).toBe("#b9cde5");
    expect(color('<a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr>')).toBe("#ff000080");
  });
});
