import { describe, expect, it } from "vite-plus/test";
import { applyContextualSpacing, splitAtPageBreaks } from "./blocks";
import type { Paragraph } from "../model";

describe("block sequence helpers", () => {
  it("removes spacing between consecutive contextual paragraphs of one style", () => {
    const first: Paragraph = {
      type: "paragraph",
      properties: { styleId: "Normal", after: 8, contextualSpacing: true },
      runs: [{ type: "text", text: "one" }],
    };
    const second: Paragraph = {
      type: "paragraph",
      properties: { styleId: "Normal", before: 6, contextualSpacing: true },
      runs: [{ type: "text", text: "two" }],
    };
    applyContextualSpacing([first, second]);
    expect(first.properties.after).toBe(0);
    expect(second.properties.before).toBe(0);
  });

  it("keeps text on either side of an explicit page break", () => {
    const segments = splitAtPageBreaks({
      type: "paragraph",
      properties: {},
      runs: [
        { type: "text", text: "A" },
        { type: "break", page: true },
        { type: "text", text: "B" },
      ],
    });
    expect(segments).toHaveLength(2);
    expect(
      segments.map((item) => item.runs.map((run) => (run.type === "text" ? run.text : ""))),
    ).toEqual([["A"], ["B"]]);
  });
});
