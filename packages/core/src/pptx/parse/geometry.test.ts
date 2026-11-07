import { describe, expect, it } from "vite-plus/test";
import { parseTransform } from "./geometry";

const xfrm = (attributes: string) =>
  new DOMParser().parseFromString(
    `<a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ${attributes}><a:off x="12700" y="25400"/><a:ext cx="127000" cy="63500"/></a:xfrm>`,
    "application/xml",
  ).documentElement;

describe("parseTransform", () => {
  it("reads offsets, extents, rotation and xsd:boolean flips", () => {
    expect(parseTransform(xfrm('rot="5400000" flipH="true" flipV="0"'))).toEqual({
      x: 1,
      y: 2,
      width: 10,
      height: 5,
      rotation: 90,
      flipH: true,
      flipV: false,
    });
    expect(parseTransform(xfrm('flipH="false" flipV="1"'))).toMatchObject({
      flipH: false,
      flipV: true,
    });
  });
});
