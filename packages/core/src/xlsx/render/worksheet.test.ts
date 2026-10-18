import { describe, expect, it } from "vite-plus/test";
import type { CellStyle, Worksheet, XlsxWorkbook } from "../model";
import { MAX_COLUMNS, MAX_ROWS } from "../layout";
import { renderWorksheet } from "./worksheet";

const style: CellStyle = {
  numberFormat: "General",
  font: { bold: true, italic: false, strike: false },
  border: {},
  alignment: { wrap: false, indent: 0 },
};

const makeSheet = (overrides: Partial<Worksheet> = {}): Worksheet => ({
  name: "Summary",
  state: "visible",
  rows: [],
  columns: [],
  merges: [],
  hyperlinks: [],
  defaultColumnWidth: 72,
  defaultRowHeight: 15,
  frozen: { rows: 0, columns: 0 },
  showGridLines: true,
  showHeaders: true,
  ...overrides,
});

const makeWorkbook = (worksheet: Worksheet): XlsxWorkbook => ({
  sheets: [worksheet],
  activeSheet: 0,
  date1904: false,
});

describe("XLSX worksheet rendering", () => {
  it("renders values, styles, safe links, headers, and merges from a model", () => {
    const worksheet = makeSheet({
      rows: [
        {
          index: 0,
          hidden: false,
          cells: [
            { column: 0, value: { type: "string", value: "Spill text" }, style },
            { column: 1, value: { type: "empty" }, style },
            {
              column: 2,
              value: { type: "number", value: 12.5 },
              style: {
                ...style,
                font: { ...style.font, bold: true, color: "#123456" },
                border: { bottom: { line: "thin", color: "#112233" } },
              },
            },
          ],
        },
        {
          index: 1,
          hidden: false,
          cells: [{ column: 0, value: { type: "string", value: "Merged" }, style }],
        },
      ],
      merges: [{ top: 1, left: 0, bottom: 1, right: 1 }],
      hyperlinks: [
        { range: { top: 0, left: 2, bottom: 0, right: 2 }, href: "https://example.com" },
      ],
      frozen: { rows: 1, columns: 1 },
    });
    const container = document.createElement("div");
    const rendered = renderWorksheet(makeWorkbook(worksheet), 0, container);

    expect(container.querySelectorAll(".orim-xlsx-column-header")).toHaveLength(3);
    expect(container.querySelectorAll(".orim-xlsx-row-header")).toHaveLength(2);
    expect(container.querySelector(".orim-xlsx-spill")?.textContent).toContain("Spill text");
    expect(container.querySelector<HTMLElement>(".orim-xlsx-cell")?.style.fontWeight).toBe("bold");
    expect(container.querySelector("a")?.getAttribute("href")).toBe("https://example.com");
    expect(container.querySelector("a")?.getAttribute("rel")).toBe("noopener noreferrer");
    expect(container.querySelector("td[colspan]")).not.toBeNull();
    rendered.destroy();
  });

  it("adds accessible resize handles only to visible column headers", () => {
    const container = document.createElement("div");
    const worksheet = makeSheet({
      rows: [
        {
          index: 0,
          hidden: false,
          cells: [{ column: 0, value: { type: "number", value: 1 }, style }],
        },
      ],
    });
    const rendered = renderWorksheet(makeWorkbook(worksheet), 0, container);
    const headers = container.querySelectorAll<HTMLElement>(".orim-xlsx-column-header");
    expect(headers.length).toBeGreaterThan(0);
    const handle = headers[0]!.querySelector<HTMLDivElement>(".orim-xlsx-resize")!;
    expect(handle.getAttribute("role")).toBe("separator");
    expect(handle.getAttribute("aria-orientation")).toBe("vertical");
    expect(handle.getAttribute("aria-label")).toBe("Resize column A");
    expect(handle.getAttribute("aria-valuemin")).toBe("8");
    const arrowRight = new KeyboardEvent("keydown", { key: "ArrowRight", cancelable: true });
    handle.dispatchEvent(arrowRight);
    expect(arrowRight.defaultPrevented).toBe(true);
    expect(handle.getAttribute("aria-valuenow")).toBe("80");
    rendered.destroy();
  });

  it("keeps resize bounds valid for zero-width columns", () => {
    const worksheet = makeSheet({
      columns: [{ min: 1, max: 1, width: 0, hidden: false }],
      rows: [
        {
          index: 0,
          hidden: false,
          cells: [{ column: 1, value: { type: "number", value: 1 }, style }],
        },
      ],
    });
    const container = document.createElement("div");
    const rendered = renderWorksheet(makeWorkbook(worksheet), 0, container);
    const handle = container.querySelectorAll<HTMLDivElement>(".orim-xlsx-resize")[1]!;

    expect(handle.getAttribute("aria-valuenow")).toBe("0");
    expect(handle.getAttribute("aria-valuemin")).toBe("0");
    rendered.destroy();
  });

  it("renders frozen cells, rich text, borders, and value-based alignment", () => {
    const formatted: CellStyle = {
      ...style,
      fill: "#eeeeee",
      font: {
        name: "Arial",
        size: 12,
        bold: true,
        italic: true,
        underline: "double",
        strike: true,
        color: "#123456",
        vertAlign: "superscript",
      },
      border: {
        top: { line: "dashed", color: "#ff0000" },
        right: { line: "double" },
        bottom: { line: "thin" },
        left: { line: "thick" },
      },
      alignment: { horizontal: "distributed", vertical: "top", wrap: true, indent: 2 },
    };
    const worksheet = makeSheet({
      frozen: { rows: 1, columns: 1 },
      rows: [
        {
          index: 0,
          hidden: false,
          cells: [
            {
              column: 0,
              value: { type: "string", value: "Styled" },
              runs: [{ text: "Run", font: { bold: true, vertAlign: "superscript" } }],
              style: formatted,
            },
            { column: 1, value: { type: "empty" }, style },
          ],
        },
        {
          index: 1,
          hidden: false,
          cells: [
            { column: 0, value: { type: "boolean", value: true }, style },
            { column: 1, value: { type: "number", value: 4 }, style },
          ],
        },
      ],
    });
    const container = document.createElement("div");
    const rendered = renderWorksheet(makeWorkbook(worksheet), 0, container);
    const styled = container.querySelector<HTMLTableCellElement>("tbody td")!;
    const cells = container.querySelectorAll<HTMLTableCellElement>("tbody td");

    expect(styled.style.position).toBe("sticky");
    expect(styled.style.zIndex).toBe("3");
    expect(styled.style.textAlign).toBe("justify");
    expect(styled.style.verticalAlign).toBe("top");
    expect(styled.style.whiteSpace).toBe("pre-wrap");
    expect(styled.style.paddingLeft).toBe("18px");
    expect(styled.style.fontWeight).toBe("bold");
    expect(styled.style.fontStyle).toBe("italic");
    expect(styled.style.textDecorationStyle).toBe("double");
    expect(styled.querySelector("span")?.textContent).toBe("Run");
    expect(styled.querySelector("span")?.style.verticalAlign).toBe("super");
    expect(styled.style.borderTop).toContain("dashed");
    expect(styled.style.borderRight).toContain("double");
    expect(cells[2]!.style.textAlign).toBe("center");
    expect(cells[3]!.style.textAlign).toBe("right");
    expect(styled.classList.contains("orim-xlsx-gridline")).toBe(false);
    expect(cells[1]!.classList.contains("orim-xlsx-gridline")).toBe(true);
    rendered.destroy();
  });

  it("renders untrusted text literally and rejects unsafe colors and links", () => {
    const worksheet = makeSheet({
      name: '<img src=x onerror="alert(1)">',
      rows: [
        {
          index: 0,
          hidden: false,
          cells: [
            {
              column: 0,
              value: { type: "string", value: '<img src=x onerror="alert(1)">' },
              style: { ...style, font: { ...style.font, color: "url(javascript:alert(1))" } },
            },
          ],
        },
      ],
      hyperlinks: [
        { range: { top: 0, left: 0, bottom: 0, right: 0 }, href: "javascript:alert(1)" },
      ],
    });
    const container = document.createElement("div");
    renderWorksheet(makeWorkbook(worksheet), 0, container);

    expect(container.querySelector("img")).toBeNull();
    expect(container.querySelector("a")).toBeNull();
    expect(container.textContent).toContain('<img src=x onerror="alert(1)">');
    expect(container.querySelector("td")?.style.color).toBe("");
  });

  it("applies row formatting only to missing cells, not present default-style cells", () => {
    const worksheet = makeSheet({
      rows: [
        {
          index: 0,
          hidden: false,
          style: { ...style, fill: "#00ff00" },
          cells: [
            { column: 0, value: { type: "number", value: 1 }, style },
            { column: 2, value: { type: "number", value: 3 }, style },
          ],
        },
      ],
    });
    const container = document.createElement("div");
    const rendered = renderWorksheet(makeWorkbook(worksheet), 0, container);
    const cells = container.querySelectorAll<HTMLElement>("tbody td");

    expect(cells).toHaveLength(3);
    expect(cells[0]!.style.backgroundColor).toBe("");
    expect(cells[1]!.style.backgroundColor).toBe("#00ff00");
    expect(cells[2]!.style.backgroundColor).toBe("");
    rendered.destroy();
  });

  it("limits sparse worksheet rendering and shows a truncation notice", () => {
    const worksheet = makeSheet({
      showHeaders: false,
      rows: [
        {
          index: MAX_ROWS + 10,
          hidden: false,
          cells: [{ column: MAX_COLUMNS + 10, value: { type: "string", value: "outside" }, style }],
        },
      ],
    });
    const container = document.createElement("div");
    const rendered = renderWorksheet(makeWorkbook(worksheet), 0, container);

    expect(container.querySelectorAll("tbody tr")).toHaveLength(100);
    expect(container.querySelectorAll("tbody td")).toHaveLength(100 * MAX_COLUMNS);
    expect(container.textContent).toContain("Showing the first 100 rows and 200 columns.");
    expect(container.textContent).not.toContain("outside");
    rendered.destroy();
  });

  it("keeps overlapping merges from occupying cells twice", () => {
    const worksheet = makeSheet({
      showHeaders: false,
      rows: [0, 1].map((index) => ({
        index,
        hidden: false,
        cells: [0, 1, 2].map((column) => ({
          column,
          value: { type: "string" as const, value: `${index}:${column}` },
          style,
        })),
      })),
      merges: [
        { top: 0, left: 0, bottom: 1, right: 1 },
        { top: 0, left: 1, bottom: 1, right: 2 },
      ],
    });
    const container = document.createElement("div");
    const rendered = renderWorksheet(makeWorkbook(worksheet), 0, container);

    expect(container.querySelectorAll("tbody tr")[0]!.querySelectorAll("td")).toHaveLength(2);
    expect(container.querySelectorAll("tbody tr")[1]!.querySelectorAll("td")).toHaveLength(1);
    expect(container.querySelector("td[rowspan='2'][colspan='2']")).not.toBeNull();
    expect(container.textContent).toContain("0:0");
    rendered.destroy();
  });
});
