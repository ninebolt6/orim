import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vite-plus/test";
import { buildDocx } from "../../test/docx";
import { buildXlsx } from "../../test/xlsx";
import { parseDocx } from "../index";
import { parseXlsx } from "./parse";

describe("XLSX package parsing", () => {
  it("follows package relationships through unusual part paths", async () => {
    const workbook = await parseXlsx(
      buildXlsx({
        workbookPath: "custom/package/workbook.xml",
        date1904: true,
        sheets: [
          {
            name: "Sales",
            path: "custom/content/grids/sales.xml",
            xml: `<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <sheetData><row r="1"><c r="A1" t="s" s="1"><v>0</v></c></row></sheetData>
              <hyperlinks><hyperlink ref="A1" r:id="rLink"/></hyperlinks>
            </worksheet>`,
            relationships: [
              {
                id: "rLink",
                type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                target: "https://example.com/sales",
                external: true,
              },
            ],
          },
        ],
        sharedStrings: {
          path: "custom/shared/strings.xml",
          xml: "<sst><si><t>Sales total</t></si></sst>",
        },
        styles: {
          path: "custom/styles/styles.xml",
          xml: `<styleSheet><fills><fill><patternFill patternType="none"/></fill>
            <fill><patternFill patternType="solid"><fgColor rgb="FF123456"/></patternFill></fill>
          </fills><cellXfs><xf fillId="0"/><xf fillId="1"/></cellXfs></styleSheet>`,
        },
        theme: {
          path: "custom/theme/theme.xml",
          xml: "<a:theme xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'/>",
        },
      }),
    );

    expect(workbook.date1904).toBe(true);
    expect(workbook.sheets[0]!.name).toBe("Sales");
    expect(workbook.sheets[0]!.rows[0]!.cells[0]!.value).toEqual({
      type: "string",
      value: "Sales total",
    });
    expect(workbook.sheets[0]!.rows[0]!.cells[0]!.style.fill).toBe("#123456");
    expect(workbook.sheets[0]!.hyperlinks).toEqual([
      { range: { top: 0, left: 0, bottom: 0, right: 0 }, href: "https://example.com/sales" },
    ]);
  });

  it("skips chartsheets while mapping the active tab to the worksheet list", async () => {
    const workbook = await parseXlsx(
      buildXlsx({
        activeTab: 1,
        sheets: [
          { name: "Chart", kind: "chartsheet", xml: "<chartsheet/>" },
          { name: "Data", xml: "<worksheet/>" },
          { name: "Hidden", state: "hidden", xml: "<worksheet/>" },
        ],
      }),
    );

    expect(workbook.sheets.map((sheet) => sheet.name)).toEqual(["Data", "Hidden"]);
    expect(workbook.activeSheet).toBe(0);
  });

  it("discovers strict-namespace relationships and ignores hyperlinks without IDs", async () => {
    const workbook = await parseXlsx(
      buildXlsx({
        strict: true,
        sheets: [
          {
            name: "Strict",
            xml: `<worksheet xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships">
              <hyperlinks><hyperlink ref="A1"/></hyperlinks>
            </worksheet>`,
          },
        ],
      }),
    );

    expect(workbook.sheets).toHaveLength(1);
    expect(workbook.sheets[0]!.hyperlinks).toEqual([]);
  });

  it("rejects invalid ZIP, empty packages, and packages without a workbook", async () => {
    await expect(parseXlsx(new Uint8Array([1, 2, 3]))).rejects.toBeInstanceOf(Error);
    await expect(parseXlsx(new Uint8Array())).rejects.toBeInstanceOf(Error);
    await expect(parseXlsx(buildXlsx({ includeWorkbook: false }))).rejects.toThrow(
      "Not an XLSX package: workbook part not found.",
    );
  });

  it("rejects DOCX packages and DOCX parsing rejects XLSX packages", async () => {
    const docx = buildDocx({ document: "<w:document/>" });
    await expect(parseXlsx(docx)).rejects.toBeInstanceOf(Error);

    const xlsx = buildXlsx({ sheets: [{ name: "Sheet", xml: "<worksheet/>" }] });
    await expect(parseDocx(xlsx)).rejects.toBeInstanceOf(Error);
  });

  it("parses small openpyxl workbooks", async () => {
    const files = [
      resolve("packages/core/test/fixtures/openpyxl-styled.xlsx"),
      resolve("packages/core/test/fixtures/openpyxl-shared-strings.xlsx"),
    ];
    const workbooks = await Promise.all(
      files.map((file) => parseXlsx(new Uint8Array(readFileSync(file)))),
    );

    expect(workbooks[0]!.sheets.map((sheet) => sheet.name)).toEqual(["Summary", "Details"]);
    expect(workbooks[0]!.sheets[0]!.rows[0]!.cells[0]!.style.fill).toBe("#1f4e78");
    expect(workbooks[0]!.sheets[0]!.merges).toContainEqual({
      top: 0,
      left: 0,
      bottom: 0,
      right: 2,
    });
    expect(workbooks[0]!.sheets[0]!.frozen).toEqual({ rows: 2, columns: 1 });
    expect(workbooks[0]!.sheets[0]!.rows[2]!.cells[0]!.style.numberFormat).toBe("yyyy-mm-dd");
    expect(workbooks[1]!.sheets[0]!.rows[1]!.cells[1]!.value).toEqual({
      type: "string",
      value: "Repeated text",
    });
  });
});
