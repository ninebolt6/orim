import { describe, expect, it } from "vite-plus/test";
import { wordElement, parseContext } from "../../../test/docx-context";
import { computeRowSpans, parseTable } from "./table";

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

  it("reads cell formatting only from tcPr and resolves grid spans and vertical merges", () => {
    const table = parseTable(
      wordElement(
        '<w:tbl><w:tr><w:tc><w:shd w:fill="FF0000"/><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>span</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>top</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>span again</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p><w:r><w:t>bottom</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
      ),
      parseContext(),
    );

    expect(table.rows[0]!.cells[0]).toMatchObject({ gridSpan: 2, shading: undefined });
    expect(table.rows[0]!.cells[1]!.rowSpan).toBe(2);
    expect(table.rows[1]!.cells[1]!.rowSpan).toBe(0);
  });
});
