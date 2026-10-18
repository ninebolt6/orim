import { describe, expect, it } from "vite-plus/test";
import { wordElement } from "../../../test/docx-context";
import { sectionProperties } from "./section";

describe("section properties parser", () => {
  it("uses default page geometry and margins", () => {
    expect(sectionProperties(undefined)).toMatchObject({
      width: 595.3,
      height: 841.9,
      orientation: "portrait",
      margins: { top: 72, right: 72, bottom: 72, left: 72, header: 36, footer: 36 },
      breakType: "nextPage",
    });
  });

  it("normalizes landscape geometry and reads explicit margins", () => {
    const properties = sectionProperties(
      wordElement(
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838" w:orient="landscape"/><w:pgMar w:top="100" w:right="200" w:bottom="300" w:left="400" w:header="500" w:footer="600"/><w:titlePg/><w:type w:val="continuous"/></w:sectPr>',
      ),
    );

    expect(properties).toMatchObject({
      width: expect.any(Number),
      height: expect.any(Number),
      orientation: "landscape",
      margins: { top: 5, right: 10, bottom: 15, left: 20, header: 25, footer: 30 },
      titlePage: true,
      breakType: "continuous",
    });
    expect(properties.width).toBeGreaterThan(properties.height);
  });

  it.each(["lines", "linesAndChars"])("uses line pitch for %s grids", (type) => {
    const properties = sectionProperties(
      wordElement(`<w:sectPr><w:docGrid w:type="${type}" w:linePitch="360"/></w:sectPr>`),
    );

    expect(properties.linePitch).toBe(18);
  });

  it("ignores line pitch for a non-line grid mode", () => {
    const properties = sectionProperties(
      wordElement('<w:sectPr><w:docGrid w:type="default" w:linePitch="360"/></w:sectPr>'),
    );

    expect(properties.linePitch).toBeUndefined();
  });
});
