import { describe, expect, it } from "vite-plus/test";
import type { Worksheet } from "./model";
import { activeSheetIndex } from "./parse";

const sheet = (name: string, state: Worksheet["state"] = "visible"): Worksheet => ({
  name,
  state,
  rows: [],
  columns: [],
  merges: [],
  hyperlinks: [],
  defaultColumnWidth: 70,
  defaultRowHeight: 15,
  frozen: { rows: 0, columns: 0 },
  showGridLines: true,
  showHeaders: true,
});

describe("active worksheet selection", () => {
  it.each([
    { label: "mapped visible tab", activeTab: 2, indexes: [[2, 1]], expected: 1 },
    { label: "hidden active tab", activeTab: 1, indexes: [[1, 0]], expected: 1 },
    { label: "invalid tab", activeTab: 8, indexes: [], expected: 1 },
    { label: "missing tab", activeTab: undefined, indexes: [], expected: 1 },
  ] as const)("selects $label", ({ activeTab, indexes, expected }) => {
    const sheets = [sheet("hidden", "hidden"), sheet("shown")];
    const mapped = activeSheetIndex(sheets, activeTab, new Map(indexes));
    expect(mapped).toBe(expected);
  });

  it("returns the first visible sheet when the active worksheet is hidden", () => {
    const sheets = [sheet("hidden", "hidden"), sheet("first"), sheet("second")];
    expect(activeSheetIndex(sheets, 0, new Map([[0, 0]]))).toBe(1);
  });

  it("returns zero when no worksheets are visible", () => {
    expect(activeSheetIndex([sheet("hidden", "hidden")], undefined, new Map())).toBe(0);
  });
});
