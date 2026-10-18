import { describe, expect, it } from "vite-plus/test";
import { xmlDocument } from "../../../test/xml";
import { defaultCellStyle, parseStyles } from "./styles";

describe("XLSX styles parser", () => {
  it("resolves font, fill, border, custom format, and alignment styles", () => {
    const styles = parseStyles(
      xmlDocument(`<styleSheet>
        <numFmts><numFmt numFmtId="166" formatCode="$#,##0.00"/></numFmts>
        <fonts><font><b/><i/><u/><sz val="12"/><name val="Arial"/><color rgb="FF123456"/></font></fonts>
        <fills><fill><patternFill patternType="solid"><fgColor rgb="FFABCDEF"/></patternFill></fill></fills>
        <borders><border><left style="medium"><color rgb="FF010203"/></left></border></borders>
        <cellXfs><xf fontId="0" fillId="0" borderId="0" numFmtId="166">
          <alignment horizontal="center" vertical="bottom" wrapText="1" indent="2"/>
        </xf></cellXfs>
      </styleSheet>`),
    );

    expect(styles.resolve(0)).toMatchObject({
      numberFormat: "$#,##0.00",
      font: {
        bold: true,
        italic: true,
        underline: "single",
        size: 12,
        name: "Arial",
        color: "#123456",
      },
      fill: "#abcdef",
      border: { left: { line: "medium", color: "#010203" } },
      alignment: { horizontal: "center", vertical: "bottom", wrap: true, indent: 2 },
    });
  });

  it("uses safe defaults for invalid style indices and CSS values", () => {
    const styles = parseStyles(
      xmlDocument(`<styleSheet>
        <fonts><font><color theme="999"/></font><font><color rgb="123"/></font>
          <font><color rgb="FF000000;background:url(x)"/></font></fonts>
        <fills><fill><patternFill patternType="solid"><fgColor indexed="9999"/></patternFill></fill>
          <fill><patternFill patternType="solid"><fgColor rgb="bad"/></patternFill></fill></fills>
        <borders><border><left style="thin"><color rgb="FF000000;background:url(x)"/></left></border></borders>
        <cellXfs><xf fontId="999" fillId="999" borderId="999" numFmtId="9999"/>
          <xf fontId="0" fillId="0" borderId="0"/><xf fontId="1"/><xf fontId="2"/></cellXfs>
      </styleSheet>`),
    );

    expect(styles.resolve(999)).toEqual(defaultCellStyle);
    expect(styles.resolve(1).font.color).toBeUndefined();
    expect(styles.resolve(2).font.color).toBeUndefined();
    expect(styles.fontForRun(undefined)).toBeUndefined();
  });

  it("validates alignment values and number formats", () => {
    const styles = parseStyles(
      xmlDocument(`<styleSheet><numFmts><numFmt numFmtId="166" formatCode="0.00"/></numFmts>
        <cellXfs><xf numFmtId="166"><alignment horizontal="center" vertical="top" wrapText="true" indent="2"/></xf>
          <xf numFmtId="9999"><alignment horizontal="unsafe" vertical="bottom" wrapText="false" indent="-1"/></xf>
        </cellXfs></styleSheet>`),
    );

    expect(styles.resolve(0)).toMatchObject({
      numberFormat: "0.00",
      alignment: { horizontal: "center", vertical: "top", wrap: true, indent: 2 },
    });
    expect(styles.resolve(1)).toMatchObject({
      numberFormat: "General",
      alignment: { vertical: "bottom", wrap: false, indent: 0 },
    });
  });
});
