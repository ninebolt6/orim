import { describe, expect, it } from "vite-plus/test";
import { border } from "./borders";
import { paragraphProperties, parseStyles } from "./styles";

describe("DOCX styles", () => {
  it("applies doc defaults, basedOn, and direct run overrides in order", () => {
    const doc = new DOMParser().parseFromString(
      `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="20"/><w:i/></w:rPr></w:rPrDefault></w:docDefaults>
        <w:style w:type="paragraph" w:styleId="Base"><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>
        <w:style w:type="paragraph" w:styleId="Child"><w:basedOn w:val="Base"/><w:rPr><w:i w:val="0"/></w:rPr></w:style>
      </w:styles>`,
      "application/xml",
    );
    const styles = parseStyles(doc);
    expect(styles.run("Child")).toMatchObject({ bold: true, italic: false, fontSize: 14 });
    expect(styles.run("Child", undefined, { bold: false })).toMatchObject({
      bold: false,
      fontSize: 14,
    });
  });

  it("resolves theme font references", () => {
    const doc = new DOMParser().parseFromString(
      `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="character" w:styleId="Theme"><w:rPr><w:rFonts w:asciiTheme="majorHAnsi"/></w:rPr></w:style></w:styles>`,
      "application/xml",
    );
    expect(
      parseStyles(doc, {
        major: "Aptos Display",
        minor: "Aptos",
        majorEastAsia: "Yu Gothic",
        minorEastAsia: "Noto Sans JP",
      }).run(undefined, "Theme").fontFamily,
    ).toBe("Aptos Display,Noto Sans JP");
    expect(
      parseStyles(
        new DOMParser().parseFromString(
          `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="character" w:styleId="EastAsia"><w:rPr><w:rFonts w:eastAsiaTheme="minorEastAsia"/></w:rPr></w:style></w:styles>`,
          "application/xml",
        ),
        { minorEastAsia: "Noto Sans JP" },
      ).run(undefined, "EastAsia").fontFamily,
    ).toBe("Noto Sans JP");
  });

  it("preserves inherited paragraph attributes and defaults a line to auto", () => {
    const doc = new DOMParser().parseFromString(
      `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Normal"><w:pPr><w:spacing w:before="100" w:after="120" w:line="300"/></w:pPr></w:style></w:styles>`,
      "application/xml",
    );
    const direct = new DOMParser().parseFromString(
      `<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:spacing w:after="0"/></w:pPr>`,
      "application/xml",
    ).documentElement;
    expect(parseStyles(doc).paragraph("Normal", paragraphProperties(direct))).toMatchObject({
      before: 5,
      after: 0,
      line: 1.25,
      lineRule: "auto",
    });
  });

  it("resolves shared borders and table-style inheritance", () => {
    const doc = new DOMParser().parseFromString(
      `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="table" w:styleId="Base"><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:color="auto"/><w:right w:val="double" w:sz="8"/></w:tblBorders><w:tblCellMar><w:left w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tcPr><w:shd w:fill="ABCDEF"/></w:tcPr></w:style><w:style w:type="table" w:default="1" w:styleId="TableGrid"><w:basedOn w:val="Base"/><w:tblPr><w:tblBorders><w:top w:val="nil"/></w:tblBorders></w:tblPr></w:style></w:styles>`,
      "application/xml",
    );
    const table = parseStyles(doc).table();
    expect(table.borders).toMatchObject({
      top: "none",
      right: expect.stringContaining("double"),
    });
    expect(table.margins?.left).toBe(5.4);
    expect(table.cellShading).toBe("#ABCDEF");
    expect(border(undefined)).toBeUndefined();
    const right = doc.getElementsByTagName("w:right")[0];
    const inheritedTop = doc.getElementsByTagName("w:top")[1];
    expect(border(right)).toContain("currentColor");
    expect(border(right)).toContain("double");
    expect(border(inheritedTop)).toBe("none");
  });
});
