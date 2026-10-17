import { describe, expect, it } from "vite-plus/test";
import "./styles.css";
import { renderWorksheet } from "./worksheet";
import type { XlsxWorkbook } from "../model";

describe("XLSX rendering in Chromium", () => {
  it("keeps headers and frozen panes sticky inside the worksheet scroller", () => {
    const workbook: XlsxWorkbook = {
      activeSheet: 0,
      date1904: false,
      sheets: [
        {
          name: "Frozen",
          state: "visible",
          rows: [
            {
              index: 0,
              hidden: false,
              cells: [
                {
                  column: 0,
                  value: { type: "string", value: "Pinned" },
                  style: {
                    numberFormat: "General",
                    font: { bold: false, italic: false, strike: false },
                    border: {},
                    alignment: { wrap: false, indent: 0 },
                  },
                },
              ],
            },
            { index: 1, hidden: false, cells: [] },
          ],
          columns: [],
          merges: [],
          hyperlinks: [],
          defaultColumnWidth: 100,
          defaultRowHeight: 20,
          frozen: { rows: 1, columns: 1 },
          showGridLines: true,
          showHeaders: true,
        },
      ],
    };
    const container = document.createElement("div");
    container.style.width = "160px";
    container.style.height = "100px";
    document.body.append(container);
    const rendered = renderWorksheet(workbook, 0, container);
    const header = container.querySelector<HTMLElement>(".orim-xlsx-header-row th")!;
    const pinned = container.querySelector<HTMLElement>("td")!;
    expect(getComputedStyle(header).position).toBe("sticky");
    expect(pinned.style.position).toBe("sticky");
    expect(pinned.style.top).toBe("24px");
    expect(pinned.style.left).toBe("48px");
    rendered.destroy();
    container.remove();
  });

  it("layers frozen headers and cells above scrolled content", async () => {
    const style = {
      numberFormat: "General",
      font: { bold: false, italic: false, strike: false },
      border: {},
      alignment: { wrap: false, indent: 0 },
    };
    const workbook: XlsxWorkbook = {
      activeSheet: 0,
      date1904: false,
      sheets: [
        {
          name: "Frozen",
          state: "visible",
          rows: Array.from({ length: 30 }, (_, rowIndex) => ({
            index: rowIndex,
            hidden: false,
            cells: Array.from({ length: 12 }, (_, column) => ({
              column,
              value: {
                type: "string" as const,
                value: `${String.fromCharCode(65 + column)}${rowIndex + 1}`,
              },
              style,
            })),
          })),
          columns: [],
          merges: [],
          hyperlinks: [],
          defaultColumnWidth: 100,
          defaultRowHeight: 24,
          frozen: { rows: 1, columns: 1 },
          showGridLines: true,
          showHeaders: true,
        },
      ],
    };
    const container = document.createElement("div");
    container.style.width = "240px";
    container.style.height = "110px";
    document.body.append(container);
    const rendered = renderWorksheet(workbook, 0, container);
    const scroll = container.querySelector<HTMLElement>(".orim-xlsx-scroll")!;
    scroll.scrollTop = scroll.scrollHeight;
    const nextFrame = (): Promise<void> =>
      new Promise((resolve) => {
        requestAnimationFrame(() => requestAnimationFrame(() => resolve()));
      });
    await nextFrame();
    const frozenRowCell = container.querySelector<HTMLElement>(
      "tbody tr:first-child td:nth-of-type(2)",
    )!;
    expect(frozenRowCell.textContent).toBe("B1");
    expect(
      document
        .elementFromPoint(
          frozenRowCell.getBoundingClientRect().left +
            frozenRowCell.getBoundingClientRect().width / 2,
          frozenRowCell.getBoundingClientRect().top +
            frozenRowCell.getBoundingClientRect().height / 2,
        )
        ?.closest("td")?.textContent,
    ).toBe("B1");

    scroll.scrollLeft = scroll.scrollWidth;
    await nextFrame();

    const rowHeader = container.querySelector<HTMLElement>(".orim-xlsx-row-header")!;
    const columnHeader = container.querySelector<HTMLElement>(".orim-xlsx-column-header")!;
    const corner = container.querySelector<HTMLElement>(".orim-xlsx-corner")!;
    const frozenCell = container.querySelector<HTMLElement>("tbody td")!;
    const elementAtCenter = (element: HTMLElement): Element | null => {
      const rect = element.getBoundingClientRect();
      return document.elementFromPoint(rect.left + rect.width / 2, rect.top + rect.height / 2);
    };

    expect(scroll.scrollTop).toBeGreaterThan(0);
    expect(scroll.scrollLeft).toBeGreaterThan(0);
    expect(elementAtCenter(rowHeader)?.closest(".orim-xlsx-row-header")?.textContent).toBe("1");
    expect(elementAtCenter(columnHeader)?.closest(".orim-xlsx-column-header")?.textContent).toBe(
      "A",
    );
    expect(elementAtCenter(corner)).toBe(corner);
    expect(elementAtCenter(frozenCell)?.closest("td")?.textContent).toBe("A1");
    rendered.destroy();
    container.remove();
  });
});
