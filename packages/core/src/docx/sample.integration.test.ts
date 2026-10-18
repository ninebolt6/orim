import { readFileSync } from "node:fs";
import { describe, expect, it } from "vite-plus/test";
import { parseDocx } from "./parse";
import type { Block, Inline } from "./model";

const collectInlines = (blocks: readonly Block[]): Inline[] => {
  return blocks.flatMap((block) =>
    block.type === "paragraph"
      ? block.runs
      : block.rows.flatMap((row) => row.cells.flatMap((cell) => collectInlines(cell.blocks))),
  );
};

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
});
