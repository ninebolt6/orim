import { describe, expect, it } from "vite-plus/test";
import type { Block, Paragraph } from "../model";
import { resolveParagraphMarks, revisionVisibility } from "./revisions";

const paragraph = (text: string, change?: "ins" | "del"): Paragraph => ({
  type: "paragraph",
  properties: change ? { change: { kind: change, author: "Aiko", date: "" } } : {},
  runs: [{ type: "text", text }],
});

describe("paragraph-mark revisions", () => {
  it("merges deleted and inserted paragraph marks into the following paragraph immutably", () => {
    const deleted = paragraph("deleted mark", "del");
    const next = paragraph("following");
    const source = Object.freeze([deleted, next]);
    Object.freeze(deleted.properties);
    Object.freeze(deleted.runs);
    Object.freeze(next.properties);
    Object.freeze(next.runs);

    const final = resolveParagraphMarks(source, "final");
    expect(final).toHaveLength(1);
    expect(final[0]).not.toBe(next);
    expect(final[0]).toMatchObject({
      properties: {},
      runs: [
        { type: "text", text: "deleted mark" },
        { type: "text", text: "following" },
      ],
    });
    expect(resolveParagraphMarks(source, "markup")).toHaveLength(2);
    expect(resolveParagraphMarks([deleted], "final")).toHaveLength(1);
    expect(
      resolveParagraphMarks([paragraph("inserted mark", "ins"), next], "original"),
    ).toHaveLength(1);
    expect(source).toHaveLength(2);
  });

  it("resolves paragraph marks recursively inside table cells", () => {
    const table: Block = {
      type: "table",
      rows: [
        {
          cells: [
            {
              gridSpan: 1,
              rowSpan: 1,
              blocks: [paragraph("removed", "del"), paragraph("kept")],
            },
          ],
        },
      ],
      widths: [],
    };
    const resolved = resolveParagraphMarks([table], "final");
    expect(resolved[0].type).toBe("table");
    if (resolved[0].type !== "table") return;
    expect(resolved[0].rows[0].cells[0].blocks).toHaveLength(1);
  });
});

describe("revision visibility", () => {
  it.each([
    ["ins", "markup", "marked"],
    ["del", "markup", "marked"],
    ["moveFrom", "markup", "marked"],
    ["moveTo", "markup", "marked"],
    ["ins", "final", "plain"],
    ["del", "final", "hidden"],
    ["moveFrom", "final", "hidden"],
    ["moveTo", "final", "plain"],
    ["ins", "original", "hidden"],
    ["del", "original", "plain"],
    ["moveFrom", "original", "plain"],
    ["moveTo", "original", "hidden"],
  ] as const)("classifies %s in %s as %s", (kind, mode, expected) => {
    expect(revisionVisibility(kind, mode)).toBe(expected);
  });
});
