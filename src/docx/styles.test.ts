import { describe, expect, it } from "vite-plus/test";
import { parseStyles } from "./styles";

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
});
