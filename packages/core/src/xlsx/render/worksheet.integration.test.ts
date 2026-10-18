import { afterEach, describe, expect, it, vi } from "vite-plus/test";
import { renderWorkbook, renderXlsx } from "./workbook";
import { renderWorksheet } from "./worksheet";
import type { Worksheet, XlsxWorkbook } from "../model";
import { buildXlsx } from "../../../test/xlsx";
import { parseXlsx } from "../parse";

const cellStyle = {
  numberFormat: "0.00",
  font: { bold: true, italic: false, strike: false, color: "#123456" },
  border: { bottom: { line: "thin" as const, color: "#112233" } },
  alignment: { wrap: false, indent: 0 },
};

const workbook: XlsxWorkbook = {
  activeSheet: 0,
  date1904: false,
  sheets: [
    {
      name: "Summary",
      state: "visible",
      rows: [
        {
          index: 0,
          hidden: false,
          cells: [
            { column: 0, value: { type: "string", value: "Spill text" }, style: cellStyle },
            { column: 1, value: { type: "empty" }, style: cellStyle },
            { column: 2, value: { type: "number", value: 12.5 }, style: cellStyle },
          ],
        },
        {
          index: 1,
          hidden: false,
          cells: [{ column: 0, value: { type: "string", value: "Merged" }, style: cellStyle }],
        },
      ],
      columns: [],
      merges: [{ top: 1, left: 0, bottom: 1, right: 1 }],
      hyperlinks: [
        { range: { top: 0, left: 2, bottom: 0, right: 2 }, href: "https://example.com" },
      ],
      defaultColumnWidth: 72,
      defaultRowHeight: 15,
      frozen: { rows: 1, columns: 1 },
      showGridLines: true,
      showHeaders: true,
    },
    {
      name: "Hidden",
      state: "hidden",
      rows: [],
      columns: [],
      merges: [],
      hyperlinks: [],
      defaultColumnWidth: 70,
      defaultRowHeight: 15,
      frozen: { rows: 0, columns: 0 },
      showGridLines: true,
      showHeaders: true,
    },
    {
      name: "Data",
      state: "visible",
      rows: [],
      columns: [],
      merges: [],
      hyperlinks: [],
      defaultColumnWidth: 70,
      defaultRowHeight: 15,
      frozen: { rows: 0, columns: 0 },
      showGridLines: true,
      showHeaders: true,
    },
  ],
};

describe("XLSX worksheet rendering", () => {
  afterEach(() => {
    document.body.replaceChildren();
    vi.restoreAllMocks();
  });

  it("renders values, rich styles, merges, safe links, headers, and sheet tabs", () => {
    const container = document.createElement("div");
    document.body.append(container);
    const onSheetChange = vi.fn();
    const rendered = renderWorkbook(workbook, container, { onSheetChange });
    expect(container.querySelector(".orim-xlsx-grid")).not.toBeNull();
    expect(container.querySelectorAll(".orim-xlsx-column-header")).toHaveLength(3);
    expect(container.querySelectorAll(".orim-xlsx-row-header")).toHaveLength(2);
    expect(container.querySelector(".orim-xlsx-spill")?.textContent).toContain("Spill text");
    expect(container.querySelector<HTMLElement>(".orim-xlsx-cell")?.style.fontWeight).toBe("bold");
    expect(container.querySelector<HTMLElement>(".orim-xlsx-cell")?.style.borderBottomWidth).toBe(
      "1px",
    );
    expect(container.querySelector("a")?.getAttribute("href")).toBe("https://example.com");
    expect(container.querySelector("a")?.getAttribute("target")).toBe("_blank");
    expect(container.querySelector("a")?.getAttribute("rel")).toBe("noopener noreferrer");
    expect(container.querySelectorAll('[role="tab"]')).toHaveLength(2);
    expect(container.querySelectorAll("td[colspan]")).toHaveLength(1);
    const nextTab = container.querySelector<HTMLButtonElement>('[role="tab"]:last-child')!;
    nextTab.click();
    expect(rendered.sheet).toBe(2);
    expect(onSheetChange).toHaveBeenCalledWith(2);
    rendered.destroy();
    rendered.update({ sheet: 0 });
    expect(container.childElementCount).toBe(0);
  });

  it("adds accessible resize handles only to visible column headers", () => {
    const container = document.createElement("div");
    const rendered = renderWorksheet(workbook, 0, container);
    const headers = container.querySelectorAll<HTMLElement>(".orim-xlsx-column-header");
    expect(headers).toHaveLength(3);
    for (const [index, header] of Array.from(headers).entries()) {
      const handles = header.querySelectorAll<HTMLDivElement>(".orim-xlsx-resize");
      expect(handles).toHaveLength(1);
      expect(handles[0]?.getAttribute("role")).toBe("separator");
      expect(handles[0]?.getAttribute("aria-orientation")).toBe("vertical");
      expect(handles[0]?.getAttribute("aria-label")).toBe(
        `Resize column ${String.fromCharCode(65 + index)}`,
      );
      expect(handles[0]?.tabIndex).toBe(0);
      expect(handles[0]?.getAttribute("aria-valuemin")).toBe("8");
      expect(handles[0]?.getAttribute("aria-valuenow")).toBe("72");
    }
    const handle = headers[1]!.querySelector<HTMLDivElement>(".orim-xlsx-resize")!;
    const column = container.querySelectorAll<HTMLTableColElement>("colgroup col")[2]!;
    const arrowRight = new KeyboardEvent("keydown", { key: "ArrowRight", cancelable: true });
    handle.dispatchEvent(arrowRight);
    expect(arrowRight.defaultPrevented).toBe(true);
    expect(column.style.width).toBe("80px");
    expect(handle.getAttribute("aria-valuenow")).toBe("80");
    for (let index = 0; index < 20; index++)
      handle.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowLeft" }));
    expect(column.style.width).toBe("8px");
    expect(handle.getAttribute("aria-valuenow")).toBe("8");
    rendered.destroy();

    const withoutHeaders: XlsxWorkbook = {
      ...workbook,
      sheets: [{ ...workbook.sheets[0]!, showHeaders: false }],
    };
    const noHeadersRender = renderWorksheet(withoutHeaders, 0, container);
    expect(container.querySelectorAll(".orim-xlsx-resize")).toHaveLength(0);
    noHeadersRender.destroy();
  });

  it("rejects hidden and out-of-range selections", () => {
    const container = document.createElement("div");
    expect(() => renderWorkbook(workbook, container, { sheet: 1 })).toThrow(RangeError);
    expect(() => renderWorksheet(workbook, 3, container)).toThrow(RangeError);
  });

  it("cancels a render lease when a different renderer reuses the container", () => {
    const container = document.createElement("div");
    const worksheet: Worksheet = workbook.sheets[0]!;
    renderWorksheet(workbook, 0, container);
    expect(container.querySelector(".orim-xlsx")).not.toBeNull();
    const next = renderWorksheet({ ...workbook, sheets: [worksheet] }, 0, container);
    expect(container.querySelectorAll(".orim-xlsx")).toHaveLength(1);
    next.destroy();
    expect(container.childElementCount).toBe(0);
  });

  it("renders untrusted text literally and rejects unsafe colors and links", () => {
    const untrusted: XlsxWorkbook = {
      ...workbook,
      sheets: [
        {
          ...workbook.sheets[0]!,
          rows: [
            {
              index: 0,
              hidden: false,
              cells: [
                {
                  column: 0,
                  value: { type: "string", value: '<img src=x onerror="alert(1)">' },
                  style: {
                    ...cellStyle,
                    font: { ...cellStyle.font, color: "url(javascript:alert(1))" },
                  },
                },
              ],
            },
          ],
          hyperlinks: [
            { range: { top: 0, left: 0, bottom: 0, right: 0 }, href: "javascript:alert(1)" },
          ],
        },
      ],
    };
    const container = document.createElement("div");
    renderWorkbook(untrusted, container);
    expect(container.querySelector("img")).toBeNull();
    expect(container.querySelector("a")).toBeNull();
    expect(container.textContent).toContain('<img src=x onerror="alert(1)">');
    expect(container.querySelector("td")?.style.color).toBe("");
  });

  it("uses a row style for a missing cell without applying it to present default-style cells", async () => {
    const parsed = await parseXlsx(
      buildXlsx({
        sheets: [
          {
            name: "Styles",
            xml: `<worksheet>
              <sheetData><row r="1" customFormat="1" s="2"><c r="A1"><v>1</v></c><c r="C1"><v>3</v></c></row></sheetData>
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
    const container = document.createElement("div");
    renderWorksheet(parsed, 0, container);
    const cells = container.querySelectorAll<HTMLElement>("tbody td");
    expect(cells).toHaveLength(3);
    expect(cells[0]!.style.backgroundColor).toBe("");
    expect(cells[1]!.style.backgroundColor).toBe("#00ff00");
    expect(cells[2]!.style.backgroundColor).toBe("");
  });

  it("rejects an already-aborted render request", async () => {
    const controller = new AbortController();
    controller.abort();
    await expect(
      renderXlsx(new Uint8Array(), document.createElement("div"), { signal: controller.signal }),
    ).rejects.toMatchObject({ name: "AbortError" });
  });

  it("cancels parsing when the caller aborts after rendering starts", async () => {
    const controller = new AbortController();
    const pending = renderXlsx(
      buildXlsx({
        sheets: [{ name: "Summary", xml: "<worksheet/>" }],
      }),
      document.createElement("div"),
      { signal: controller.signal },
    );
    controller.abort();
    await expect(pending).rejects.toMatchObject({ name: "AbortError" });
  });
});
