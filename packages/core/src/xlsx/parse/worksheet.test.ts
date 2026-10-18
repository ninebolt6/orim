import { describe, expect, it } from "vite-plus/test";
import { xmlDocument, xmlRoot } from "../../../test/xml";
import { parseStyles } from "./styles";
import { parseSharedStrings, parseWorksheet } from "./worksheet";

const parse = (xml: string, overrides: Partial<Parameters<typeof parseWorksheet>[1]> = {}) => {
  return parseWorksheet(xmlRoot(xml), {
    name: "Sheet",
    state: "visible",
    date1904: false,
    shared: [],
    styles: parseStyles(),
    relationships: new Map(),
    ...overrides,
  });
};

describe("XLSX worksheet parser", () => {
  it("ignores invalid references, deduplicates cells, and orders rows and cells", () => {
    const sheet = parse(`<worksheet><sheetData>
      <row r="0"><c r="A1"><v>9</v></c></row>
      <row r="-2"><c r="A1"><v>8</v></c></row>
      <row r="1e9"><c r="A1"><v>7</v></c></row>
      <row r="3"><c r="C3"><v>3</v></c><c r="A3"><v>1</v></c><c r="A3"><v>2</v></c></row>
      <row r="1"><c r="A0"><v>1</v></c><c r="A1048577"><v>2</v></c><c r="XFE1"><v>3</v></c>
        <c r="1A"><v>4</v></c><c r="A1"><v>1</v></c><c r=""><v>5</v></c>
        <c r="garbage"><v>6</v></c><c r="C1"><v>3</v></c><c r="A1"><v>2</v></c>
      </row>
    </sheetData></worksheet>`);

    expect(sheet.rows.map((row) => row.index)).toEqual([0, 2]);
    expect(sheet.rows[0]!.cells.map((cell) => [cell.column, cell.value])).toEqual([
      [0, { type: "number", value: 2 }],
      [1, { type: "number", value: 5 }],
      [2, { type: "number", value: 3 }],
    ]);
    expect(sheet.rows[1]!.cells.map((cell) => [cell.column, cell.value])).toEqual([
      [0, { type: "number", value: 2 }],
      [2, { type: "number", value: 3 }],
    ]);
  });

  it("parses cached values, rich inline text, errors, and invalid shared-string indices", () => {
    const styles = parseStyles();
    const shared = parseSharedStrings(
      xmlRoot("<sst><si><t>Plain </t><r><rPr><b/></rPr><t>Bold</t></r></si></sst>"),
      styles,
    );
    const sheet = parse(
      `<worksheet><sheetData><row r="1">
        <c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>999</v></c>
        <c r="C1"><f>1+1</f><v>3</v></c><c r="D1"><f>1+1</f></c>
        <c r="E1" t="e"><v>#DIV/0!</v></c>
        <c r="F1" t="inlineStr"><is><t>before </t><r><rPr><b/></rPr><t>after</t></r></is></c>
      </row></sheetData></worksheet>`,
      { shared },
    );

    expect(sheet.rows[0]!.cells.map((cell) => cell.value)).toEqual([
      { type: "string", value: "Plain Bold" },
      { type: "empty" },
      { type: "number", value: 3 },
      { type: "empty" },
      { type: "error", value: "#DIV/0!" },
      { type: "string", value: "before after" },
    ]);
    expect(sheet.rows[0]!.cells[0]!.runs).toMatchObject([
      { text: "Plain " },
      { text: "Bold", font: { bold: true } },
    ]);
    expect(parseSharedStrings(undefined, styles)).toEqual([]);
  });

  it("parses remaining cell value types and date systems", () => {
    const sheet = parse(
      `<worksheet><sheetData><row r="1">
        <c r="A1" t="str"><v>_x0041_</v></c><c r="B1" t="b"><v>1</v></c>
        <c r="C1" t="b"><v>0</v></c><c r="D1" t="b"><v>2</v></c>
        <c r="E1" t="d"><v>2020-01-01T00:00:00Z</v></c><c r="F1" t="d"><v>invalid</v></c>
        <c r="G1"><v>NaN</v></c><c r="H1" t="unknown"><v>1</v></c>
        <c r="I1" t="inlineStr"><is/></c><c r="J1" t="s"><v>-1</v></c>
      </row></sheetData></worksheet>`,
    );
    const date1904 = parse(
      `<worksheet><sheetData><row r="1"><c r="A1" t="d"><v>1904-01-02</v></c></row></sheetData></worksheet>`,
      { date1904: true },
    );

    expect(sheet.rows[0]!.cells.map((cell) => cell.value)).toEqual([
      { type: "string", value: "A" },
      { type: "boolean", value: true },
      { type: "boolean", value: false },
      { type: "empty" },
      { type: "number", value: 43831 },
      { type: "empty" },
      { type: "empty" },
      { type: "empty" },
      { type: "empty" },
      { type: "empty" },
    ]);
    expect(date1904.rows[0]!.cells[0]!.value).toEqual({ type: "number", value: 1 });
  });

  it("applies the default cell style rather than a row style to unstyled cells", () => {
    const styles = parseStyles(
      xmlDocument(`<styleSheet><fills>
        <fill><patternFill patternType="none"/></fill>
        <fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill>
        <fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/></patternFill></fill>
      </fills><cellXfs><xf fillId="0"/><xf fillId="1"/><xf fillId="2"/></cellXfs></styleSheet>`),
    );
    const sheet = parse(
      `<worksheet><cols><col min="2" max="2" style="1"/></cols>
        <sheetData><row r="1" customFormat="1" s="2">
          <c r="A1" s="1"><v>1</v></c><c r="B1"><v>2</v></c><c r="C1"><v>3</v></c>
        </row></sheetData></worksheet>`,
      { styles },
    );

    expect(sheet.rows[0]!.style?.fill).toBe("#00ff00");
    expect(sheet.rows[0]!.cells.map((cell) => cell.style.fill)).toEqual([
      "#ff0000",
      undefined,
      undefined,
    ]);
    expect(sheet.columns[0]!.style?.fill).toBe("#ff0000");
  });

  it("drops unsafe and malformed worksheet hyperlinks", () => {
    const relationships = new Map([
      [
        "safe",
        { id: "safe", type: "rel/hyperlink", target: "https://example.com", external: true },
      ],
      [
        "unsafe",
        { id: "unsafe", type: "rel/hyperlink", target: "javascript:alert(1)", external: true },
      ],
      ["internal", { id: "internal", type: "rel/worksheet", target: "sheet.xml", external: false }],
    ]);
    const sheet = parse(
      `<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <hyperlinks><hyperlink ref="A1" r:id="safe"/><hyperlink ref="B1" r:id="unsafe"/>
          <hyperlink ref="C1" r:id="internal"/><hyperlink ref="D1"/>
          <hyperlink ref="E1" foo:id="safe"/><hyperlink ref="garbage" r:id="safe"/>
        </hyperlinks>
      </worksheet>`,
      { relationships },
    );

    expect(sheet.hyperlinks).toEqual([
      { range: { top: 0, left: 0, bottom: 0, right: 0 }, href: "https://example.com" },
    ]);
  });

  it("clamps dimensions and frozen panes while omitting invalid rows and columns", () => {
    const sheet = parse(`<worksheet>
      <sheetFormatPr defaultColWidth="1e9" defaultRowHeight="1e9"/>
      <cols><col min="1" max="1" width="1e9"/><col min="4" max="2" width="20"/>
        <col min="5" max="999999" width="20"/><col min="2" max="2" width="NaN"/>
        <col min="3" max="3" width="-1"/></cols>
      <sheetViews><sheetView><pane state="frozen" xSplit="1e9" ySplit="-1"/></sheetView></sheetViews>
      <sheetData><row r="1" ht="1e9"><c r="A1"><v>1</v></c></row>
        <row r="2" ht="NaN"><c r="B2"><v>2</v></c></row>
        <row r="1048577"><c r="A1048577"><v>3</v></c></row>
      </sheetData>
    </worksheet>`);

    expect(sheet.defaultColumnWidth).toBeLessThanOrEqual(1785);
    expect(sheet.defaultRowHeight).toBe(409);
    expect(sheet.rows.map((row) => row.height)).toEqual([409, undefined]);
    expect(sheet.rows).toHaveLength(2);
    expect(sheet.columns).toHaveLength(3);
    expect(sheet.frozen).toEqual({ rows: 0, columns: 200 });
  });
});
