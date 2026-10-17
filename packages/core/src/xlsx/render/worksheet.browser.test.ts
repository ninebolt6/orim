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
});
