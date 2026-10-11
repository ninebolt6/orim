import { describe, expect, it } from "vite-plus/test";
import { computeRowSpans } from "./table";

describe("table row spans", () => {
  it("matches vertical-merge continuations by grid column across grid spans", () => {
    const rows = Object.freeze([
      Object.freeze({
        cells: Object.freeze([
          Object.freeze({ gridSpan: 2, rowSpan: 1 }),
          Object.freeze({ gridSpan: 1, rowSpan: -1 }),
        ]),
      }),
      Object.freeze({
        cells: Object.freeze([
          Object.freeze({ gridSpan: 2, rowSpan: 1 }),
          Object.freeze({ gridSpan: 1, rowSpan: 0 }),
        ]),
      }),
    ]);

    expect(computeRowSpans(rows)).toEqual([
      [1, 2],
      [1, 0],
    ]);
    expect(rows[0].cells[1].rowSpan).toBe(-1);
  });
});
