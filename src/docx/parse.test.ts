import { readFileSync } from "node:fs";
import { describe, expect, it } from "vite-plus/test";
import { parseDocx } from "../render/render";
import type { Block, Inline } from "./model";

function blocks(document: Awaited<ReturnType<typeof parseDocx>>): Block[] {
  return document.sections.flatMap((section) => [
    ...section.blocks,
    ...(section.properties.header ?? []),
    ...(section.properties.footer ?? []),
  ]);
}

function inlines(blockList: Block[]): Inline[] {
  return blockList.flatMap((block) =>
    block.type === "paragraph"
      ? block.runs
      : block.rows.flatMap((row) => row.cells.flatMap((cell) => inlines(cell.blocks))),
  );
}

describe("DOCX parser", () => {
  it("retains cross-paragraph comments, replies, changes, merged cells, and fields", async () => {
    const data = new Uint8Array(readFileSync("fixtures/sample.docx"));
    const doc = await parseDocx(data);
    const all = blocks(doc);
    const runs = inlines(all);
    const anchored = runs.filter((run) => run.type === "text" && run.comments?.includes("0"));
    expect(anchored.length).toBeGreaterThanOrEqual(2);
    expect(doc.comments.find((comment) => comment.id === "1")).toMatchObject({
      parentId: "0",
      done: true,
    });
    expect(doc.sections).toHaveLength(2);
    expect(doc.sections.map((section) => section.properties.orientation)).toEqual([
      "portrait",
      "landscape",
    ]);
    expect(doc.imageParts["word/media/sample.png"]).toBeDefined();
    expect(runs.some((run) => run.type === "text" && run.href === "https://example.com/orim")).toBe(
      true,
    );
    expect(
      runs.flatMap((run) => (run.type === "text" && run.change ? [run.change.kind] : [])),
    ).toEqual(expect.arrayContaining(["ins", "del", "moveFrom", "moveTo"]));
    expect(runs.some((run) => run.type === "text" && run.change?.author === "Aiko")).toBe(true);
    expect(runs.some((run) => run.type === "text" && run.formattingChange?.author === "Aiko")).toBe(
      true,
    );
    const table = all.find((block) => block.type === "table");
    expect(table?.type === "table" && table.rows[1].cells[0].rowSpan).toBe(2);
    expect(table?.type === "table" && table.rows.some((row) => row.change?.kind === "ins")).toBe(
      true,
    );
    expect(runs.filter((run) => run.type === "field").map((run) => run.name)).toEqual(
      expect.arrayContaining(["PAGE", "NUMPAGES"]),
    );
  });
});
