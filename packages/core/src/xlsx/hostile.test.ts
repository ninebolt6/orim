import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { createElement } from "react";
import { unzipSync, zipSync } from "fflate";
import { cleanup, render } from "@testing-library/react";
import { afterEach, describe, expect, it } from "vite-plus/test";
import { parseDocx } from "../index";
import { XlsxViewer } from "../../../react/src/XlsxViewer";
import { buildXlsx } from "../../test/xlsx";
import { formatCellValue } from "./format";
import { MAX_COLUMNS, MAX_ROWS } from "./layout";
import { parseXlsx } from "./parse";
import { renderWorkbook } from "./render/workbook";

const timeout = 2000;
const plainSheet = (xml: string) => buildXlsx({ sheets: [{ name: "Hostile", xml }] });

afterEach(() => cleanup());

describe("hostile XLSX packages and XML", () => {
  it(
    "rejects non-ZIP, truncated, empty, and unrelated packages with errors",
    async () => {
      for (const bytes of [
        new Uint8Array([1, 2, 3]),
        new Uint8Array([0x50, 0x4b, 0x03, 0x04]),
        zipSync({}),
        buildXlsx({ includeWorkbook: false }),
      ]) {
        await expect(parseXlsx(bytes)).rejects.toBeInstanceOf(Error);
      }
      await expect(parseXlsx(buildXlsx({ includeWorkbook: false }))).rejects.toThrow(
        "Not an XLSX package: workbook part not found.",
      );
    },
    timeout,
  );

  it(
    "enforces the ZIP uncompressed-entry limit and rejects cross-format packages",
    async () => {
      const archive = unzipSync(plainSheet("<worksheet/>"));
      const compressed = zipSync(archive, { level: 9 });
      const bytes = new Uint8Array(compressed);
      const view = new DataView(bytes.buffer);
      for (let offset = 0; offset + 46 < bytes.length; offset++) {
        if (view.getUint32(offset, true) !== 0x02014b50) continue;
        const nameLength = view.getUint16(offset + 28, true);
        const name = new TextDecoder().decode(
          bytes.subarray(offset + 46, offset + 46 + nameLength),
        );
        if (!name.endsWith("workbook.xml")) continue;
        view.setUint32(offset + 24, 64 * 1024 * 1024 + 1, true);
        break;
      }
      await expect(parseXlsx(bytes)).rejects.toThrow("ZIP entry exceeds the 64 MiB limit");
      const docx = new Uint8Array(readFileSync(resolve("fixtures/sample.docx")));
      await expect(parseXlsx(docx)).rejects.toBeInstanceOf(Error);
      await expect(parseDocx(plainSheet("<worksheet/>"))).rejects.toBeInstanceOf(Error);
    },
    timeout,
  );

  it(
    "rejects malformed XML and bounds internal-entity expansion",
    async () => {
      await expect(parseXlsx(plainSheet("<worksheet><sheetData>"))).rejects.toBeInstanceOf(Error);
      const declarations = ['<!ENTITY e0 "x">'];
      for (let index = 1; index < 10; index++)
        declarations.push(
          `<!ENTITY e${index} "${Array.from({ length: 10 }, () => `&e${index - 1};`).join("")}">`,
        );
      const doctype = (root: string, inner: string) =>
        `<!DOCTYPE ${root} [${declarations.join("")}]>${inner}`;
      const cases = [
        buildXlsx({
          sharedStrings: {
            xml: doctype("sst", "<sst><si><t>&e7;</t></si></sst>"),
          },
          sheets: [
            {
              name: "Entity",
              xml: '<worksheet><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>',
            },
          ],
        }),
        plainSheet(
          doctype(
            "worksheet",
            '<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>',
          ),
        ),
      ];
      for (const bytes of cases) {
        let parsed;
        try {
          parsed = await parseXlsx(bytes);
        } catch (error) {
          expect(error).toBeInstanceOf(Error);
          continue;
        }
        const text = parsed.sheets[0]?.rows[0]?.cells[0]?.value;
        expect(text?.type === "string" ? text.value.length : 0).toBeLessThan(1_000_000);
      }
    },
    timeout,
  );
});

describe("hostile XLSX references, indices, styles, and values", () => {
  it(
    "ignores invalid references and safely deduplicates and orders cells",
    async () => {
      const parsed = await parseXlsx(
        plainSheet(`<worksheet><sheetData>
          <row r="0"><c r="A1"><v>9</v></c></row>
          <row r="-2"><c r="A1"><v>8</v></c></row>
          <row r="1e9"><c r="A1"><v>7</v></c></row>
          <row r="1">
            <c r="A0"><v>1</v></c><c r="A1048577"><v>2</v></c><c r="XFE1"><v>3</v></c>
            <c r="1A"><v>4</v></c><c r="A1"><v>1</v></c><c r=""><v>5</v></c>
            <c r="garbage"><v>6</v></c><c r="C1"><v>3</v></c><c r="A1"><v>2</v></c>
          </row>
        </sheetData></worksheet>`),
      );
      expect(parsed.sheets[0]!.rows.map((row) => row.index)).toEqual([0]);
      expect(parsed.sheets[0]!.rows[0]!.cells.map((cell) => [cell.column, cell.value])).toEqual([
        [0, { type: "number", value: 2 }],
        [1, { type: "number", value: 5 }],
        [2, { type: "number", value: 3 }],
      ]);
    },
    timeout,
  );

  it(
    "defaults invalid shared-string, style, format, font, fill, border, theme, and palette indices",
    async () => {
      const parsed = await parseXlsx(
        buildXlsx({
          sharedStrings: { xml: "<sst><si><t>safe</t></si></sst>" },
          theme: {
            xml: '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="x"/></a:themeElements></a:theme>',
          },
          styles: {
            xml: `<styleSheet>
              <fonts><font><color theme="999"/></font><font><color rgb="123"/></font>
                <font><color rgb="123456789"/></font><font><color rgb="FF123456" tint="2"/></font>
                <font><color rgb="FF123456" tint="-2"/></font></fonts>
              <fills><fill><patternFill patternType="solid"><fgColor indexed="9999"/></patternFill></fill></fills>
              <borders><border><left style="thin"><color rgb="FF000000;background:url(x)"/></left></border></borders>
              <cellXfs><xf fontId="999" fillId="999" borderId="999" numFmtId="9999"/>
                <xf fontId="1"/><xf fontId="2"/><xf fontId="3"/><xf fontId="4"/></cellXfs>
            </styleSheet>`,
          },
          sheets: [
            {
              name: "Indices",
              xml: `<worksheet><sheetData><row r="1">
                <c r="A1" t="s"><v>999</v></c><c r="B1" t="s"><v>-1</v></c>
                <c r="C1" t="s"><v>NaN</v></c><c r="D1" s="999"><v>2</v></c>
                <c r="E1" s="1"><v>2</v></c><c r="F1" s="2"><v>2</v></c>
                <c r="G1" s="3"><v>2</v></c><c r="H1" s="4"><v>2</v></c>
              </row></sheetData></worksheet>`,
            },
          ],
        }),
      );
      const cells = parsed.sheets[0]!.rows[0]!.cells;
      expect(cells.slice(0, 3).map((cell) => cell.value)).toEqual([
        { type: "empty" },
        { type: "empty" },
        { type: "empty" },
      ]);
      expect(cells[3]!.style.numberFormat).toBe("General");
      expect(cells[3]!.style.fill).toBeUndefined();
      expect(cells[3]!.style.border).toEqual({});
      expect(cells.slice(4).map((cell) => cell.style.font.color)).toEqual([
        undefined,
        undefined,
        "#ffffff",
        "#000000",
      ]);
    },
    timeout,
  );

  it(
    "keeps hostile number formats and numeric extremes bounded and non-throwing",
    () => {
      const formats = [
        "0".repeat(10_000),
        '"'.repeat(10_000),
        "[".repeat(10_000),
        "[>]",
        "[Red",
        "0".repeat(1000),
        "#,##0,,,,,,,,,",
        "E+",
        "?/??????????",
        "[h]:mm",
        "@;@;@;@;@",
        ";;;",
      ];
      const values = [NaN, Infinity, -Infinity, 1e308, -0, 5e-324, -1e100, 1e10];
      for (const format of formats)
        for (const value of values)
          expect(typeof formatCellValue({ type: "number", value }, format, false).text).toBe(
            "string",
          );
      for (const value of values)
        for (const format of ["[h]:mm", "yyyy-mm-dd"])
          expect(typeof formatCellValue({ type: "number", value }, format, false).text).toBe(
            "string",
          );
      expect(formatCellValue({ type: "number", value: 1e10 }, "[h]:mm", false).text).toBe(
        "########",
      );
      for (const value of [-1e100, 1e100])
        expect(formatCellValue({ type: "number", value }, "yyyy-mm-dd", false).text).toBe(
          "########",
        );
    },
    timeout,
  );

  it(
    "sanitizes hostile links and emits safe rel attributes for accepted links",
    async () => {
      const targets = [
        "javascript:alert(1)",
        "JaVaScRiPt:alert(1)",
        "java\tscript:alert(1)",
        " \tjavascript:alert(1)",
        "data:text/html,boom",
        "vbscript:alert(1)",
        "file:///etc/passwd",
        "relative/path",
        "https://safe.example/",
        "#location",
      ];
      const relationships = targets.map((target, index) => ({
        id: `r${index}`,
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        target,
        external: true,
      }));
      const links = targets
        .map((_, index) => `<hyperlink ref="A${index + 1}" r:id="r${index}"/>`)
        .join("");
      const bytes = buildXlsx({
        sheets: [
          {
            name: "Links",
            relationships,
            xml: `<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1">${targets.map((_, index) => `<c r="A${index + 1}" t="inlineStr"><is><t>x</t></is></c>`).join("")}</row></sheetData><hyperlinks>${links}</hyperlinks></worksheet>`,
          },
        ],
      });
      const parsed = await parseXlsx(bytes);
      const container = document.createElement("div");
      renderWorkbook(parsed, container);
      const anchors = [...container.querySelectorAll("a")];
      expect(anchors.map((anchor) => anchor.getAttribute("href"))).toEqual([
        "https://safe.example/",
        "#location",
      ]);
      expect(anchors.every((anchor) => anchor.rel === "noopener noreferrer")).toBe(true);
      expect(parsed.sheets[0]!.hyperlinks).toHaveLength(2);
    },
    timeout,
  );

  it(
    "renders hostile sheet and cell text as text and keeps CSS values to hex colors",
    async () => {
      const parsed = await parseXlsx(
        buildXlsx({
          sharedStrings: { xml: "<sst><si><t>&lt;img src=x onerror=alert(1)&gt;</t></si></sst>" },
          styles: {
            xml: `<styleSheet><fonts><font><color rgb="FF000000;background:url(x)"/></font></fonts>
              <fills><fill><patternFill patternType="solid"><fgColor rgb="zzz"/></patternFill></fill></fills>
              <cellXfs><xf fontId="0" fillId="0"/></cellXfs></styleSheet>`,
          },
          sheets: [
            {
              name: "&lt;img src=x onerror=alert(1)&gt;",
              xml: `<worksheet><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="inlineStr"><is><t>&lt;/table&gt;&lt;script&gt;</t></is></c></row></sheetData></worksheet>`,
            },
          ],
        }),
      );
      const container = document.createElement("div");
      renderWorkbook(parsed, container);
      expect(container.querySelectorAll("img,script")).toHaveLength(0);
      expect(container.textContent).toContain("<img src=x onerror=alert(1)>");
      expect(container.textContent).toContain("</table><script>");
      for (const element of container.querySelectorAll<HTMLElement>("[style]")) {
        const style = element.getAttribute("style")!;
        expect(style).not.toMatch(/url\s*\(|javascript:|NaN/i);
        for (const declaration of style.split(";")) {
          const [property, value] = declaration.split(":");
          if (property?.trim().endsWith("color")) expect(value?.trim()).toMatch(/^#[0-9a-f]{6}$/i);
        }
        for (const color of style.match(/#[0-9a-f]{3,9}/gi) ?? [])
          expect(color).toMatch(/^#[0-9a-f]{6}$/i);
      }
    },
    timeout,
  );

  it(
    "clamps hostile row, column, and freeze dimensions to finite values",
    async () => {
      const dimensions = `<worksheet><sheetFormatPr defaultColWidth="1e9" defaultRowHeight="1e9"/>
        <cols><col min="1" max="1" width="1e9"/><col min="4" max="2" width="20"/>
          <col min="5" max="999999" width="20"/><col min="2" max="2" width="NaN"/>
          <col min="3" max="3" width="-1"/></cols>
        <sheetData><row r="1" ht="1e9"><c r="A1"><v>1</v></c></row><row r="2" ht="NaN"><c r="B2"><v>2</v></c></row></sheetData>
      </worksheet>`;
      for (const [xSplit, ySplit] of [
        ["1e9", "-1"],
        ["-1", "1e9"],
        ["NaN", "NaN"],
      ]) {
        const parsed = await parseXlsx(
          plainSheet(
            dimensions.replace(
              "<sheetData>",
              `<sheetViews><sheetView><pane state="frozen" xSplit="${xSplit}" ySplit="${ySplit}"/></sheetView></sheetViews><sheetData>`,
            ),
          ),
        );
        const sheet = parsed.sheets[0]!;
        expect(Number.isFinite(sheet.defaultColumnWidth)).toBe(true);
        expect(Number.isFinite(sheet.defaultRowHeight)).toBe(true);
        expect(sheet.defaultColumnWidth).toBeLessThanOrEqual(1785);
        expect(sheet.defaultRowHeight).toBeLessThanOrEqual(409);
        expect(sheet.columns.every((column) => (column.width ?? 0) <= 1785)).toBe(true);
        expect(sheet.rows.every((row) => row.height === undefined || row.height <= 409)).toBe(true);
        expect(sheet.frozen.columns).toBeLessThanOrEqual(MAX_COLUMNS);
        expect(sheet.frozen.rows).toBeLessThanOrEqual(MAX_ROWS);
        expect(Number.isFinite(sheet.frozen.columns)).toBe(true);
        expect(Number.isFinite(sheet.frozen.rows)).toBe(true);
      }
    },
    timeout,
  );
});

describe("hostile XLSX workbook edge cases", () => {
  it(
    "handles empty, hidden-only, invalid-active-tab, and chartsheet-only workbooks",
    async () => {
      const cases: readonly { readonly bytes: Uint8Array; readonly hasVisibleSheet: boolean }[] = [
        { bytes: buildXlsx(), hasVisibleSheet: false },
        {
          bytes: buildXlsx({ sheets: [{ name: "Hidden", state: "hidden", xml: "<worksheet/>" }] }),
          hasVisibleSheet: false,
        },
        {
          bytes: buildXlsx({
            activeTab: -10,
            sheets: [{ name: "Visible", xml: "<worksheet/>" }],
          }),
          hasVisibleSheet: true,
        },
        {
          bytes: buildXlsx({
            activeTab: 999,
            sheets: [{ name: "Visible", xml: "<worksheet/>" }],
          }),
          hasVisibleSheet: true,
        },
        {
          bytes: buildXlsx({
            sheets: [{ name: "Chart", kind: "chartsheet", xml: "<chartsheet/>" }],
          }),
          hasVisibleSheet: false,
        },
      ];
      for (const { bytes, hasVisibleSheet } of cases) {
        const parsed = await parseXlsx(bytes);
        expect(parsed.activeSheet).toBe(0);
        const container = document.createElement("div");
        const rendered = renderWorkbook(parsed, container);
        if (hasVisibleSheet) expect(container.querySelector(".orim-xlsx-grid")).not.toBeNull();
        else expect(container.textContent).toContain("No visible worksheets.");
        rendered.destroy();
        const view = render(createElement(XlsxViewer, { workbook: parsed }));
        expect(view.container.querySelectorAll("img,script")).toHaveLength(0);
        expect(view.container.querySelector(".orim-xlsx-grid") !== null).toBe(hasVisibleSheet);
      }
    },
    timeout,
  );

  it(
    "uses cached formula values and displays error cells as text",
    async () => {
      const parsed = await parseXlsx(
        plainSheet(`<worksheet><sheetData><row r="1">
          <c r="A1"><f>1+1</f><v>3</v></c><c r="B1"><f>1+1</f></c>
          <c r="C1" t="e"><v>#DIV/0!</v></c>
        </row></sheetData></worksheet>`),
      );
      expect(parsed.sheets[0]!.rows[0]!.cells.map((cell) => cell.value)).toEqual([
        { type: "number", value: 3 },
        { type: "empty" },
        { type: "error", value: "#DIV/0!" },
      ]);
      const container = document.createElement("div");
      renderWorkbook(parsed, container);
      expect(container.textContent).toContain("3");
      expect(container.textContent).toContain("#DIV/0!");
    },
    timeout,
  );
});

describe("real-world XLSX interoperability", () => {
  it(
    "parses and renders small openpyxl workbooks with styles, merges, freezes, dates, and shared strings",
    async () => {
      const files = [
        resolve("packages/core/test/fixtures/openpyxl-styled.xlsx"),
        resolve("packages/core/test/fixtures/openpyxl-shared-strings.xlsx"),
      ];
      for (const file of files) {
        const parsed = await parseXlsx(new Uint8Array(readFileSync(file)));
        const container = document.createElement("div");
        expect(() => renderWorkbook(parsed, container)).not.toThrow();
        expect(container.querySelector(".orim-xlsx-grid")).not.toBeNull();
        expect(parsed.sheets.length).toBe(2);
        if (file.endsWith("openpyxl-styled.xlsx")) {
          const summary = parsed.sheets[0]!;
          expect(summary.name).toBe("Summary");
          expect(summary.rows[0]!.cells[0]!.style.font.bold).toBe(true);
          expect(summary.rows[0]!.cells[0]!.style.fill).toBe("#1f4e78");
          expect(summary.merges).toContainEqual({ top: 0, left: 0, bottom: 0, right: 2 });
          expect(summary.frozen).toEqual({ rows: 2, columns: 1 });
          expect(summary.rows[2]!.cells[0]!.style.numberFormat).toBe("yyyy-mm-dd");
          expect(summary.rows[2]!.cells[2]!.value).toEqual({ type: "string", value: "Ready" });
          expect(parsed.sheets[1]!.name).toBe("Details");
        } else {
          const strings = parsed.sheets[0]!;
          expect(strings.name).toBe("Strings");
          expect(strings.rows[1]!.cells[0]!.value).toEqual({ type: "string", value: "Item 0" });
          expect(strings.rows[1]!.cells[1]!.value).toEqual({
            type: "string",
            value: "Repeated text",
          });
          expect(parsed.sheets[1]!.name).toBe("Dates");
        }
      }
    },
    timeout,
  );
});
