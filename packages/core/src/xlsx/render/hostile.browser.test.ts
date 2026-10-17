import { describe, expect, it } from "vite-plus/test";
import { buildXlsx } from "../../../test/xlsx";
import { MAX_COLUMNS, MAX_ROWS } from "../layout";
import { parseXlsx } from "../parse";
import { renderWorksheet } from "./worksheet";
import type { XlsxWorkbook } from "../model";
import "./styles.css";

const timeout = 2000;

describe("hostile XLSX layout in Chromium", () => {
  it(
    "renders a sparse maximum-coordinate sheet within the layout caps",
    async () => {
      const workbook = await parseXlsx(
        buildXlsx({
          sheets: [
            {
              name: "Sparse",
              xml: `<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row>
                <row r="1048576"><c r="XFD1048576"><v>2</v></c></row></sheetData></worksheet>`,
            },
          ],
        }),
      );
      const container = document.createElement("div");
      document.body.append(container);
      const rendered = renderWorksheet(workbook, 0, container);
      const grid = container.querySelector(".orim-xlsx-grid")!;
      expect(grid.querySelectorAll("tbody tr").length).toBeLessThanOrEqual(MAX_ROWS);
      expect(grid.querySelectorAll("tbody td").length).toBeLessThanOrEqual(20_000);
      expect(grid.querySelectorAll("thead tr").length).toBe(1);
      expect(grid.querySelectorAll("thead th").length).toBeLessThanOrEqual(MAX_COLUMNS + 1);
      expect(container.textContent).toMatch(/Showing the first \d+ rows and \d+ columns\./);
      rendered.destroy();
      container.remove();
    },
    timeout,
  );

  it(
    "normalizes reversed merges, ignores overlapping and malformed merges, and keeps row spans aligned",
    async () => {
      const parsed = await parseXlsx(
        buildXlsx({
          sheets: [
            {
              name: "Merges",
              xml: `<worksheet><sheetData>
                <row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c><c r="C1"><v>3</v></c></row>
                <row r="2"><c r="A2"><v>4</v></c><c r="B2"><v>5</v></c><c r="C2"><v>6</v></c></row>
                <row r="3"><c r="A3"><v>7</v></c><c r="B3"><v>8</v></c><c r="C3"><v>9</v></c></row>
              </sheetData><mergeCells>
                <mergeCell ref="B2:A1"/><mergeCell ref="B1:C2"/><mergeCell ref="C3:C3"/>
                <mergeCell ref="XFD1048577:A1"/><mergeCell ref="garbage"/>
              </mergeCells></worksheet>`,
            },
          ],
        }),
      );
      const container = document.createElement("div");
      document.body.append(container);
      const rendered = renderWorksheet(parsed, 0, container);
      const rows = [...container.querySelectorAll<HTMLTableRowElement>("tbody tr")];
      expect(rows).toHaveLength(3);
      const occupancy: boolean[][] = Array.from({ length: rows.length }, () => []);
      rows.forEach((row, rowIndex) => {
        for (const cell of row.querySelectorAll<HTMLTableCellElement>("td")) {
          let columnIndex = occupancy[rowIndex]!.findIndex((occupied) => !occupied);
          if (columnIndex < 0) columnIndex = occupancy[rowIndex]!.length;
          for (let rowOffset = 0; rowOffset < cell.rowSpan; rowOffset++)
            for (let columnOffset = 0; columnOffset < cell.colSpan; columnOffset++)
              occupancy[rowIndex + rowOffset]![columnIndex + columnOffset] = true;
        }
      });
      expect(occupancy.map((row) => row.slice(0, 3).filter(Boolean).length)).toEqual([3, 3, 3]);
      expect(occupancy.every((row) => row.length <= 3)).toBe(true);
      rendered.destroy();
      container.remove();
    },
    timeout,
  );

  it(
    "renders clamped hostile dimensions without NaN style values",
    async () => {
      const parsed: XlsxWorkbook = await parseXlsx(
        buildXlsx({
          sheets: [
            {
              name: "Dimensions",
              xml: `<worksheet><cols><col min="1" max="1" width="1e9"/></cols>
                <sheetViews><sheetView><pane state="frozen" xSplit="1e9" ySplit="1e9"/></sheetView></sheetViews>
                <sheetData><row r="1" ht="1e9"><c r="A1"><v>1</v></c></row></sheetData></worksheet>`,
            },
          ],
        }),
      );
      const container = document.createElement("div");
      document.body.append(container);
      const rendered = renderWorksheet(parsed, 0, container);
      for (const element of container.querySelectorAll<HTMLElement>("[style]"))
        expect(element.getAttribute("style")).not.toMatch(/NaN|Infinity/i);
      rendered.destroy();
      container.remove();
    },
    timeout,
  );
});
