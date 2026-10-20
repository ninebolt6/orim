import { describe, expect, it } from "vite-plus/test";
import type { Worksheet } from "./model";
import { computeLayout, HEADER_HEIGHT, HEADER_WIDTH, MAX_CELLS, visibleSpan } from "./layout";

const worksheet = (overrides: Partial<Worksheet> = {}): Worksheet => ({
  name: "Sheet",
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
  ...overrides,
});

const style = (fill?: string) => ({
  numberFormat: "General",
  font: { bold: false, italic: false, strike: false },
  ...(fill ? { fill } : {}),
  border: {},
  alignment: { wrap: false, indent: 0 },
});

describe("XLSX layout", () => {
  it("finds visible positions intersecting an inclusive index range", () => {
    const items = [{ index: 2 }, { index: 4 }, { index: 7 }, { index: 10 }];
    expect(visibleSpan([], 0, 10)).toEqual([0, 0]);
    expect(visibleSpan(items, 0, 1)).toEqual([0, 0]);
    expect(visibleSpan(items, 3, 7)).toEqual([1, 3]);
    expect(visibleSpan(items, 11, 12)).toEqual([4, 4]);
    expect(visibleSpan([{ index: 0 }, { index: 2 }, { index: 5 }, { index: 9 }], 1, 8)).toEqual([
      1, 3,
    ]);
  });

  it("derives extent from content, styled cells, merges, and frozen panes", () => {
    const layout = computeLayout(
      worksheet({
        rows: [
          {
            index: 0,
            hidden: false,
            cells: [{ column: 2, value: { type: "empty" }, style: style("#fff000") }],
          },
          { index: 4, hidden: false, style: style("#000000"), cells: [] },
        ],
        columns: [{ min: 8, max: 9, hidden: false, style: style("#000000") }],
        merges: [{ top: 5, left: 4, bottom: 6, right: 6 }],
        frozen: { rows: 2, columns: 3 },
      }),
    );
    expect(layout.rows.map((row) => row.index)).toEqual([0, 1, 2, 3, 4, 5, 6]);
    expect(layout.columns.map((column) => column.index)).toEqual([0, 1, 2, 3, 4, 5, 6, 7, 8, 9]);
    expect(layout.rows[0]!.stickyTop).toBe(HEADER_HEIGHT);
    expect(layout.columns[2]!.stickyLeft).toBe(HEADER_WIDTH + 140);
  });

  it("adds a minimal opposite dimension for styled-only extents", () => {
    const styledRow = computeLayout(
      worksheet({
        rows: [{ index: 2, hidden: false, style: style("#000000"), cells: [] }],
      }),
    );
    expect(styledRow.rows.map((row) => row.index)).toEqual([0, 1, 2]);
    expect(styledRow.columns.map((column) => column.index)).toEqual([0]);

    const styledColumns = computeLayout(
      worksheet({
        columns: [{ min: 0, max: 1, hidden: false, style: style("#000000") }],
      }),
    );
    expect(styledColumns.columns.map((column) => column.index)).toEqual([0, 1]);
    expect(styledColumns.rows.map((row) => row.index)).toEqual([0]);

    const fullColumn = computeLayout(
      worksheet({
        columns: [{ min: 0, max: 16383, hidden: false, style: style("#000000") }],
      }),
    );
    expect(fullColumn.rows).toHaveLength(0);
    expect(fullColumn.columns).toHaveLength(0);
  });

  it("omits hidden rows and columns and anchors merges at the first visible position", () => {
    const layout = computeLayout(
      worksheet({
        rows: [
          {
            index: 0,
            hidden: false,
            cells: [{ column: 0, value: { type: "string", value: "anchor" }, style: style() }],
          },
          { index: 1, hidden: true, cells: [] },
        ],
        columns: [
          { min: 0, max: 0, hidden: true },
          { min: 1, max: 1, hidden: false },
        ],
        merges: [{ top: 0, left: 0, bottom: 1, right: 1 }],
      }),
    );
    expect(layout.rows.map((row) => row.index)).toEqual([0]);
    expect(layout.columns.map((column) => column.index)).toEqual([1]);
    expect(layout.merges).toMatchObject([
      {
        anchorRow: 0,
        anchorColumn: 1,
        rowSpan: 1,
        columnSpan: 1,
        cell: { value: { value: "anchor" } },
      },
    ]);
  });

  it("caps visible dimensions and marks truncation", () => {
    const layout = computeLayout(
      worksheet({
        rows: [
          {
            index: 5000,
            hidden: false,
            cells: [{ column: 205, value: { type: "number", value: 1 }, style: style() }],
          },
        ],
      }),
    );
    expect(layout.rows).toHaveLength(100);
    expect(layout.columns).toHaveLength(200);
    expect(layout.truncated).toBe(true);
  });

  it("bounds merge coverage by the cell cap", () => {
    const layout = computeLayout(
      worksheet({
        merges: [{ top: 0, left: 0, bottom: 4999, right: 199 }],
      }),
    );
    expect(layout.rows).toHaveLength(100);
    expect(layout.columns).toHaveLength(200);
    expect(layout.covered.size).toBeLessThan(MAX_CELLS);
    expect(layout.truncated).toBe(true);
  });

  it("short-circuits duplicate full-sheet merges", () => {
    const merge = { top: 0, left: 0, bottom: 4999, right: 199 };
    const start = performance.now();
    const layout = computeLayout(
      worksheet({
        merges: Array.from({ length: 10_000 }, () => merge),
      }),
    );
    expect(layout.merges).toHaveLength(1);
    expect(performance.now() - start).toBeLessThan(2000);
  }, 2000);

  it("anchors frozen cells at the origin when headers are hidden", () => {
    const layout = computeLayout(
      worksheet({
        showHeaders: false,
        rows: [{ index: 0, hidden: false, cells: [] }],
        columns: [{ min: 0, max: 0, hidden: false }],
        frozen: { rows: 1, columns: 1 },
      }),
    );
    expect(layout.rows[0]!.stickyTop).toBe(0);
    expect(layout.columns[0]!.stickyLeft).toBe(0);
  });
});
