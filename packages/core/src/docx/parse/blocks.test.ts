import { describe, expect, it } from "vite-plus/test";
import { children } from "../xml";
import { wordElement } from "../../../test/docx-context";
import { applyContextualSpacing, blockNodes, splitAtPageBreaks } from "./blocks";
import type { Paragraph } from "../model";

describe("block sequence helpers", () => {
  it("removes spacing between consecutive contextual paragraphs of one style", () => {
    const first: Paragraph = Object.freeze({
      type: "paragraph",
      properties: Object.freeze({ styleId: "Normal", after: 8, contextualSpacing: true }),
      runs: Object.freeze([{ type: "text" as const, text: "one" }]),
    });
    const second: Paragraph = Object.freeze({
      type: "paragraph",
      properties: Object.freeze({ styleId: "Normal", before: 6, contextualSpacing: true }),
      runs: Object.freeze([{ type: "text" as const, text: "two" }]),
    });
    const input = Object.freeze([first, second]);
    const output = applyContextualSpacing(input);
    expect(input[0].properties.after).toBe(8);
    expect(input[1].properties.before).toBe(6);
    expect(output[0].type === "paragraph" && output[0].properties.after).toBe(0);
    expect(output[1].type === "paragraph" && output[1].properties.before).toBe(0);
    expect(output[0]).not.toBe(first);
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

  it("unwraps nested block-level content controls", () => {
    const body = wordElement(
      "<w:body><w:sdt><w:sdtContent><w:p><w:r><w:t>TOC entry</w:t></w:r></w:p></w:sdtContent></w:sdt></w:body>",
    );

    expect(blockNodes(children(body)).map((node) => node.localName)).toEqual(["p"]);
  });
});
