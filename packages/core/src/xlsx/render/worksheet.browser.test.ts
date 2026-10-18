import { describe, expect, it } from "vite-plus/test";
import "./styles.css";
import { renderWorksheet } from "./worksheet";
import type { XlsxWorkbook } from "../model";

const resizeWorkbook = (frozenColumns: number, columnCount: number): XlsxWorkbook => {
  const style = {
    numberFormat: "General",
    font: { bold: false, italic: false, strike: false },
    border: {},
    alignment: { wrap: false, indent: 0 },
  };
  return {
    activeSheet: 0,
    date1904: false,
    sheets: [
      {
        name: "Resize",
        state: "visible",
        rows: Array.from({ length: 8 }, (_, rowIndex) => ({
          index: rowIndex,
          hidden: false,
          cells: Array.from({ length: columnCount }, (_, column) => ({
            column,
            value: { type: "string" as const, value: `cell-${rowIndex}-${column}` },
            style,
          })),
        })),
        columns: [],
        merges: [],
        hyperlinks: [],
        defaultColumnWidth: 100,
        defaultRowHeight: 24,
        frozen: { rows: 0, columns: frozenColumns },
        showGridLines: true,
        showHeaders: true,
      },
    ],
  };
};

const mockPointerCapture = (handle: HTMLElement): void => {
  let capturedPointer: number | undefined;
  handle.setPointerCapture = (pointerId) => {
    capturedPointer = pointerId;
  };
  handle.hasPointerCapture = (pointerId) => capturedPointer === pointerId;
  handle.releasePointerCapture = (pointerId) => {
    if (capturedPointer === pointerId) capturedPointer = undefined;
  };
};

const dispatchPointer = (
  handle: HTMLElement,
  type: string,
  pointerId: number,
  clientX: number,
): void => {
  handle.dispatchEvent(
    new PointerEvent(type, {
      bubbles: true,
      cancelable: true,
      clientX,
      pointerId,
      pointerType: "mouse",
    }),
  );
};

describe("XLSX rendering in Chromium", () => {
  it("shows header edges when worksheet gridlines are disabled", () => {
    const workbook: XlsxWorkbook = {
      activeSheet: 0,
      date1904: false,
      sheets: [
        {
          name: "No gridlines",
          state: "visible",
          rows: [
            {
              index: 0,
              hidden: false,
              cells: [
                {
                  column: 0,
                  value: { type: "string", value: "value" },
                  style: {
                    numberFormat: "General",
                    font: { bold: false, italic: false, strike: false },
                    border: {},
                    alignment: { wrap: false, indent: 0 },
                  },
                },
              ],
            },
          ],
          columns: [],
          merges: [],
          hyperlinks: [],
          defaultColumnWidth: 100,
          defaultRowHeight: 24,
          frozen: { rows: 0, columns: 0 },
          showGridLines: false,
          showHeaders: true,
        },
      ],
    };
    const container = document.createElement("div");
    document.body.append(container);
    const rendered = renderWorksheet(workbook, 0, container);
    const columnHeader = container.querySelector<HTMLElement>(".orim-xlsx-column-header")!;
    const rowHeader = container.querySelector<HTMLElement>(".orim-xlsx-row-header")!;
    const corner = container.querySelector<HTMLElement>(".orim-xlsx-corner")!;
    expect(getComputedStyle(columnHeader).boxShadow).not.toBe("none");
    expect(getComputedStyle(rowHeader).boxShadow).not.toBe("none");
    expect(getComputedStyle(corner).boxShadow).not.toBe("none");
    rendered.destroy();
    container.remove();
  });

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

  it("resizes columns by pointer and enforces the minimum width", () => {
    const container = document.createElement("div");
    container.style.width = "600px";
    container.style.height = "300px";
    document.body.append(container);
    const rendered = renderWorksheet(resizeWorkbook(0, 8), 0, container);
    const table = container.querySelector<HTMLTableElement>("table")!;
    const column = container.querySelectorAll<HTMLTableColElement>("colgroup col")[2]!;
    const header = container.querySelectorAll<HTMLElement>(".orim-xlsx-column-header")[1]!;
    const handle = header.querySelector<HTMLElement>(".orim-xlsx-resize")!;
    const initialColumnWidth = Number.parseFloat(column.style.width);
    const initialHeaderWidth = header.getBoundingClientRect().width;
    const initialTableWidth = table.getBoundingClientRect().width;
    mockPointerCapture(handle);
    const pointerDown = new PointerEvent("pointerdown", {
      bubbles: true,
      cancelable: true,
      clientX: 100,
      pointerId: 1,
      pointerType: "mouse",
    });
    handle.dispatchEvent(pointerDown);
    expect(pointerDown.defaultPrevented).toBe(true);
    dispatchPointer(handle, "pointermove", 1, 150);
    dispatchPointer(handle, "pointerup", 1, 150);
    expect(Number.parseFloat(column.style.width)).toBe(initialColumnWidth + 50);
    expect(header.getBoundingClientRect().width).toBeCloseTo(initialHeaderWidth + 50, 0);
    expect(table.getBoundingClientRect().width).toBeCloseTo(initialTableWidth + 50, 0);

    dispatchPointer(handle, "pointerdown", 2, 200);
    dispatchPointer(handle, "pointermove", 2, -800);
    dispatchPointer(handle, "pointerup", 2, -800);
    expect(Number.parseFloat(column.style.width)).toBe(8);
    rendered.destroy();
    container.remove();
  });

  it("updates frozen column offsets when resizing and keeps them pinned on scroll", async () => {
    const container = document.createElement("div");
    container.style.width = "240px";
    container.style.height = "140px";
    document.body.append(container);
    const rendered = renderWorksheet(resizeWorkbook(2, 8), 0, container);
    const scroll = container.querySelector<HTMLElement>(".orim-xlsx-scroll")!;
    const headers = container.querySelectorAll<HTMLElement>(".orim-xlsx-column-header");
    const columnAHandle = headers[0]!.querySelector<HTMLElement>(".orim-xlsx-resize")!;
    const headerB = headers[1]!;
    const cellB = container.querySelector<HTMLElement>("tbody tr:first-child td:nth-of-type(2)")!;
    mockPointerCapture(columnAHandle);
    dispatchPointer(columnAHandle, "pointerdown", 3, 100);
    dispatchPointer(columnAHandle, "pointermove", 3, 130);
    dispatchPointer(columnAHandle, "pointerup", 3, 130);
    expect(headerB.style.left).toBe("178px");
    expect(cellB.style.left).toBe("178px");

    scroll.scrollLeft = 500;
    await new Promise<void>((resolve) => {
      requestAnimationFrame(() => requestAnimationFrame(() => resolve()));
    });
    const scrollLeft = scroll.getBoundingClientRect().left;
    expect(scroll.scrollLeft).toBeGreaterThan(0);
    expect(headerB.getBoundingClientRect().left).toBeCloseTo(scrollLeft + 178, 0);
    expect(cellB.getBoundingClientRect().left).toBeCloseTo(scrollLeft + 178, 0);
    rendered.destroy();
    container.remove();
  });
});
