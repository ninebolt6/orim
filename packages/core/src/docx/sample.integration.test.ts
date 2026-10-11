import { readFileSync } from "node:fs";
import { describe, expect, it } from "vite-plus/test";
import { parseDocx } from "./parse";
import type { Block, Inline } from "./model";
import { renderDocument } from "../render/document";

function collectInlines(blocks: readonly Block[]): Inline[] {
  return blocks.flatMap((block) =>
    block.type === "paragraph"
      ? block.runs
      : block.rows.flatMap((row) => row.cells.flatMap((cell) => collectInlines(cell.blocks))),
  );
}

describe("sample DOCX model", () => {
  it("retains TOC links and footer page fields", async () => {
    const doc = await parseDocx(new Uint8Array(readFileSync("fixtures/sample.docx")));
    const main = collectInlines(doc.sections[0].blocks);
    expect(
      main.some(
        (run) =>
          run.type === "text" && run.text.includes("Overview") && run.href === "#_TocOverview",
      ),
    ).toBe(true);
    expect(main.some((run) => run.type === "text" && run.text.includes("Second TOC entry"))).toBe(
      true,
    );
    const footer = collectInlines(doc.sections[0].properties.footer ?? []);
    expect(footer.filter((run) => run.type === "field").map((run) => run.name)).toEqual(
      expect.arrayContaining(["PAGE", "NUMPAGES"]),
    );
  });

  it("parses and renders the sample comment reply", async () => {
    const doc = await parseDocx(new Uint8Array(readFileSync("fixtures/sample.docx")));
    const aikoComment = doc.comments.find((comment) => comment.author === "Aiko");
    expect(aikoComment).toBeDefined();
    const reply = doc.comments.find((comment) => comment.author === "Morgan");
    expect(reply?.parentId).toBe(aikoComment!.id);

    const container = document.createElement("div");
    const rendered = renderDocument(doc, container);
    const replyBoxes = container.querySelectorAll(".orim-comment-replies");
    expect(replyBoxes).toHaveLength(1);
    expect(replyBoxes[0].textContent).toContain("Agreed — the wording is clear.");
    rendered.destroy();
  });
});
