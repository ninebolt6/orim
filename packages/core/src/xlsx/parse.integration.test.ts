import { describe, expect, it } from "vite-plus/test";
import { buildXlsx } from "../../test/xlsx";
import { parseXlsx } from "./parse";

const relationshipBase = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

describe("XLSX parser integration", () => {
  it("discovers parts through relationships and parses sparse worksheet content", async () => {
    const workbook = await parseXlsx(
      buildXlsx({
        workbookPath: "irregular/location/book-main.xml",
        activeTab: 1,
        date1904: true,
        sheets: [
          {
            name: "Chart",
            kind: "chartsheet",
            xml: "<chartsheet/>",
          },
          {
            name: "Data",
            path: "separate/sheets/data-table.xml",
            xml: `<worksheet xmlns:r="${relationshipBase}">
              <sheetViews><sheetView showGridLines="0" showRowColHeaders="0"><pane xSplit="2" ySplit="1" state="frozen"/></sheetView></sheetViews>
              <sheetFormatPr defaultColWidth="12" defaultRowHeight="18"/>
              <cols><col min="2" max="4" width="10" style="0"/><col min="5" max="5" hidden="1"/></cols>
              <sheetData>
                <row><c><v>12.5</v></c><c t="s"><v>0</v></c><c t="inlineStr"><is><r><rPr><b/><color rgb="FF123456"/></rPr><t>inline</t></r><rPh><t>phonetic</t></rPh><r><t> text</t></r></is></c></row>
                <row r="3" ht="21" hidden="1" customFormat="1" s="0"><c r="$B$3" s="0"><f>A1+B1</f><v>45000</v></c><c t="b"><v>1</v></c><c t="b"><v>0</v></c><c t="e"><v>#DIV/0!</v></c><c t="str"><v>_x005F_x0041__x000D_</v></c><c t="d"><v>1904-01-02T12:00:00Z</v></c><c t="n"><v>NaN</v></c><c t="s"><v>99</v></c></row>
              </sheetData>
              <mergeCells><mergeCell ref="$A$1:B2"/><mergeCell ref="invalid"/></mergeCells>
              <hyperlinks><hyperlink ref="A1" r:id="safe"/><hyperlink ref="B1" r:id="unsafe"/><hyperlink ref="C1" location="Summary!A1"/><hyperlink ref="D1"/></hyperlinks>
            </worksheet>`,
            relationships: [
              {
                id: "safe",
                type: `${relationshipBase}/hyperlink`,
                target: "https://example.com/sheet",
                external: true,
              },
              {
                id: "unsafe",
                type: `${relationshipBase}/hyperlink`,
                target: "javascript:alert(1)",
                external: true,
              },
            ],
          },
        ],
        sharedStrings: {
          path: "odd/strings/shared-table.xml",
          xml: `<sst>
            <si><t>_x005F_x0041__x000D_</t></si>
            <si><rPh><t>phonetic</t></rPh><r><rPr><b/><rFont val="Aptos"/></rPr><t>Rich</t></r><r><t> text</t></r></si>
          </sst>`,
        },
        theme: {
          path: "themes/unusual/theme.xml",
          xml: `<a:theme xmlns:a="urn:theme"><a:themeElements><a:clrScheme><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2><a:dk2><a:srgbClr val="111111"/></a:dk2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme></a:themeElements></a:theme>`,
        },
        styles: {
          path: "formatting/cell-styles.xml",
          xml: `<styleSheet>
            <numFmts><numFmt numFmtId="166" formatCode="0.00"/></numFmts>
            <fonts><font><b/><i val="0"/><name val="Aptos"/><sz val="12"/><color theme="4"/><u/><strike val="0"/></font></fonts>
            <fills><fill><patternFill patternType="solid"><fgColor rgb="FFAABBCC"/></patternFill></fill></fills>
            <borders><border><left style="thin"><color rgb="FF334455"/></left></border></borders>
            <cellXfs><xf numFmtId="166" fontId="0" fillId="0" borderId="0"><alignment horizontal="center" wrapText="1" indent="2"/></xf></cellXfs>
          </styleSheet>`,
        },
      }),
    );
    expect(workbook).toMatchObject({
      activeSheet: 0,
      date1904: true,
      sheets: [
        {
          name: "Data",
          state: "visible",
          defaultColumnWidth: 84,
          defaultRowHeight: 18,
          frozen: { rows: 1, columns: 2 },
          showGridLines: false,
          showHeaders: false,
          columns: [
            { min: 1, max: 3, width: 70, hidden: false },
            { min: 4, max: 4, hidden: true },
          ],
          merges: [{ top: 0, left: 0, bottom: 1, right: 1 }],
          hyperlinks: [
            { range: { top: 0, left: 0, bottom: 0, right: 0 }, href: "https://example.com/sheet" },
          ],
          rows: [
            {
              index: 0,
              cells: [
                { column: 0, value: { type: "number", value: 12.5 } },
                { column: 1, value: { type: "string", value: "_x0041_\r" } },
                { column: 2, value: { type: "string", value: "inline text" } },
              ],
            },
            {
              index: 2,
              height: 21,
              hidden: true,
              cells: [
                { column: 1, value: { type: "number", value: 45000 } },
                { column: 2, value: { type: "boolean", value: true } },
                { column: 3, value: { type: "boolean", value: false } },
                { column: 4, value: { type: "error", value: "#DIV/0!" } },
                { column: 5, value: { type: "string", value: "_x0041_\r" } },
                { column: 6, value: { type: "number", value: 1.5 } },
                { column: 7, value: { type: "empty" } },
                { column: 8, value: { type: "empty" } },
              ],
            },
          ],
        },
      ],
    });
    const data = workbook.sheets[0]!;
    expect(data.rows[0]!.cells[1]!.runs).toBeUndefined();
    expect(data.rows[0]!.cells[2]!.runs?.map((run) => run.text)).toEqual(["inline", " text"]);
    expect(data.rows[0]!.cells[2]!.runs?.[0]!.font).toMatchObject({
      bold: true,
      color: "#123456",
    });
    expect(data.rows[1]!.cells[0]!.style).toMatchObject({
      numberFormat: "0.00",
      font: {
        name: "Aptos",
        size: 12,
        bold: true,
        italic: false,
        underline: "single",
        strike: false,
        color: "#4472c4",
      },
      fill: "#aabbcc",
      border: { left: { line: "thin", color: "#334455" } },
      alignment: { horizontal: "center", wrap: true, indent: 2 },
    });
    expect(data.rows[1]!.cells[1]!.style).toBe(data.rows[1]!.cells[0]!.style);
  });

  it("maps active worksheet tabs, skips chartsheets, and uses visible-sheet fallback", async () => {
    const workbook = await parseXlsx(
      buildXlsx({
        activeTab: 0,
        sheets: [
          { name: "Chart", kind: "chartsheet", xml: "<chartsheet/>" },
          { name: "Hidden", state: "veryHidden", xml: "<worksheet/>" },
          { name: "Visible", xml: "<worksheet/>" },
        ],
      }),
    );
    expect(workbook.sheets.map((sheet) => [sheet.name, sheet.state])).toEqual([
      ["Hidden", "veryHidden"],
      ["Visible", "visible"],
    ]);
    expect(workbook.activeSheet).toBe(1);
  });

  it("throws the specified error when the workbook relationship is absent", async () => {
    await expect(parseXlsx(buildXlsx({ includeWorkbook: false }))).rejects.toThrow(
      "Not an XLSX package: workbook part not found.",
    );
  });

  it("discovers strict namespace relationships and ignores links without relationship ids", async () => {
    const workbook = await parseXlsx(
      buildXlsx({
        strict: true,
        sheets: [
          {
            name: "Strict",
            xml: `<worksheet>
              <sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>value</t></is></c></row></sheetData>
              <hyperlinks><hyperlink ref="A1"/></hyperlinks>
            </worksheet>`,
          },
        ],
      }),
    );
    expect(workbook.sheets[0]!.rows[0]!.cells[0]!.value).toEqual({
      type: "string",
      value: "value",
    });
    expect(workbook.sheets[0]!.hyperlinks).toEqual([]);
  });

  it("rejects invalid ZIP packages", async () => {
    await expect(parseXlsx(new Uint8Array([1, 2, 3]))).rejects.toThrow();
  });

  it("uses the default cell style when a present cell omits its style ID", async () => {
    const workbook = await parseXlsx(
      buildXlsx({
        sheets: [
          {
            name: "Styles",
            xml: `<worksheet>
              <cols><col min="2" max="2" style="1"/></cols>
              <sheetData>
                <row r="1" customFormat="1" s="2"><c r="A1" s="3"><v>1</v></c><c r="B1"><v>2</v></c></row>
                <row r="2"><c r="B2"><v>3</v></c><c r="C2"><v>4</v></c></row>
              </sheetData>
            </worksheet>`,
          },
        ],
        styles: {
          xml: `<styleSheet>
            <fills>
              <fill><patternFill patternType="none"/></fill>
              <fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill>
              <fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/></patternFill></fill>
              <fill><patternFill patternType="solid"><fgColor rgb="FF0000FF"/></patternFill></fill>
            </fills>
            <cellXfs>
              <xf fillId="0"/><xf fillId="1"/><xf fillId="2"/><xf fillId="3"/>
            </cellXfs>
          </styleSheet>`,
        },
      }),
    );
    const [firstRow, secondRow] = workbook.sheets[0]!.rows;
    expect(firstRow!.cells.map((cell) => cell.style.fill)).toEqual(["#0000ff", undefined]);
    expect(secondRow!.cells.map((cell) => cell.style.fill)).toEqual([undefined, undefined]);
  });

  it("uses cellXfs[0] for an unstyled cell even when the row has a style", async () => {
    const workbook = await parseXlsx(
      buildXlsx({
        sheets: [
          {
            name: "Styles",
            xml: `<worksheet>
              <sheetData><row r="1" customFormat="1" s="2"><c r="A1"><v>1</v></c></row></sheetData>
            </worksheet>`,
          },
        ],
        styles: {
          xml: `<styleSheet>
            <fills>
              <fill><patternFill patternType="none"/></fill>
              <fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill>
              <fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/></patternFill></fill>
            </fills>
            <cellXfs><xf fillId="0"/><xf fillId="1"/><xf fillId="2"/></cellXfs>
          </styleSheet>`,
        },
      }),
    );
    const row = workbook.sheets[0]!.rows[0]!;
    expect(row.style?.fill).toBe("#00ff00");
    expect(row.cells[0]!.style.fill).toBeUndefined();
  });
});
