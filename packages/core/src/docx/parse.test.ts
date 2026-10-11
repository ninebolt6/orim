import { readFileSync } from "node:fs";
import { describe, expect, it } from "vite-plus/test";
import { parseDocx } from "./parse";
import { splitAtPageBreaks } from "./parse/blocks";
import type { Block, Inline, Paragraph } from "./model";

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
    expect(runs.some((run) => run.type === "text" && run.text.includes("Overview"))).toBe(true);
    expect(runs.some((run) => run.type === "text" && run.href === "#_TocOverview")).toBe(true);
    expect(
      runs.flatMap((run) => (run.type === "text" && run.change ? [run.change.kind] : [])),
    ).toEqual(expect.arrayContaining(["ins", "del", "moveFrom", "moveTo"]));
    expect(runs.some((run) => run.type === "text" && run.change?.author === "Aiko")).toBe(true);
    expect(runs.some((run) => run.type === "text" && run.formattingChange?.author === "Aiko")).toBe(
      true,
    );
    const table = all.find((block) => block.type === "table");
    expect(table?.type === "table" && table.width).toBe("100%");
    expect(table?.type === "table" && table.borders?.top).toContain("currentColor");
    expect(table?.type === "table" && table.rows[1].cells[0].shading).toBe("#F6F8FC");
    expect(table?.type === "table" && table.rows[1].cells[1].rowSpan).toBe(2);
    expect(table?.type === "table" && table.rows[1].cells[0].gridSpan).toBe(2);
    expect(table?.type === "table" && table.rows.some((row) => row.change?.kind === "ins")).toBe(
      true,
    );
    expect(runs.filter((run) => run.type === "field").map((run) => run.name)).toEqual(
      expect.arrayContaining(["PAGE", "NUMPAGES"]),
    );
    const pageFields = runs.filter((run) => run.type === "field" && run.name === "PAGE");
    expect(pageFields).toHaveLength(3);
    expect(
      pageFields.map((run) => (run.type === "field" ? run.style?.fontSize : undefined)).sort(),
    ).toEqual([10, 10, 15]);
    expect(doc.sections[0].properties.linePitch).toBe(18);
    expect(doc.defaultTabStop).toBe(36);
    const inheritedNormal = all.find(
      (block) =>
        block.type === "paragraph" &&
        block.runs.some((run) => run.type === "text" && run.text.includes("日本語の文章")),
    );
    expect(inheritedNormal?.type === "paragraph" && inheritedNormal.properties).toMatchObject({
      before: 5,
      after: 6,
      line: 1.25,
      lineRule: "auto",
    });
    expect(all.some((block) => block.type === "paragraph" && block.list)).toBe(true);
    expect(
      all.some(
        (block) =>
          block.type === "paragraph" &&
          block.runs.some(
            (run) => run.type === "text" && run.text.includes("Numbering cancelled"),
          ) &&
          !block.list,
      ),
    ).toBe(true);
    const styleNumbered = all.find(
      (block) =>
        block.type === "paragraph" &&
        block.runs.some(
          (run) => run.type === "text" && run.text.includes("numbered heading inherited"),
        ),
    );
    const styleLevelOverride = all.find(
      (block) =>
        block.type === "paragraph" &&
        block.runs.some((run) => run.type === "text" && run.text.includes("Direct list level")),
    );
    expect(styleNumbered?.type === "paragraph" && styleNumbered.list).toBeDefined();
    expect(styleLevelOverride?.type === "paragraph" && styleLevelOverride.list?.level).toBe(1);
    const directZero = all.find(
      (block) =>
        block.type === "paragraph" &&
        block.runs.some((run) => run.type === "text" && run.text.includes("direct zero after")),
    );
    expect(directZero?.type === "paragraph" && directZero.properties).toMatchObject({
      before: 5,
      after: 0,
      line: 1.25,
      lineRule: "auto",
    });
    const contextual = all.filter(
      (block) =>
        block.type === "paragraph" &&
        block.runs.some((run) => run.type === "text" && run.text.includes("contextual-spacing")),
    );
    expect(contextual[0]?.type === "paragraph" && contextual[0].properties.after).toBe(0);
    expect(contextual[1]?.type === "paragraph" && contextual[1].properties.before).toBe(0);
    const before = all.find(
      (block) =>
        block.type === "paragraph" &&
        block.runs.some(
          (run) => run.type === "text" && run.text.includes("Before inline page break"),
        ),
    );
    const after = all.find(
      (block) =>
        block.type === "paragraph" &&
        block.runs.some(
          (run) => run.type === "text" && run.text.includes("After inline page break"),
        ),
    );
    expect(before?.type === "paragraph" && before.properties.pageBreakAfter).toBe(true);
    expect(after?.type === "paragraph" && after.properties.before).toBe(0);
    expect(after?.type === "paragraph" && after.list).toBeUndefined();
    expect(
      all.some(
        (block) =>
          block.type === "paragraph" &&
          block.runs.some((run) => run.type === "text" && run.text.includes("Second TOC entry")),
      ),
    ).toBe(true);
    expect(
      all.some(
        (block) =>
          block.type === "paragraph" &&
          block.runs.some((run) => run.type === "text" && run.text.includes("Contents")),
      ),
    ).toBe(true);
    expect(
      all.some(
        (block) =>
          block.type === "paragraph" &&
          block.runs.some((run) => run.type === "text" && run.text.includes("Overview")),
      ),
    ).toBe(true);
    expect(runs.some((run) => run.type === "text" && run.href === "#_TocOverview")).toBe(true);
    expect(runs.some((run) => run.type === "text" && run.href === "https://example.org/orim")).toBe(
      true,
    );
    const borderParagraph = all.find(
      (block) =>
        block.type === "paragraph" &&
        block.runs.some((run) => run.type === "text" && run.text.includes("bottom border")),
    );
    expect(
      borderParagraph?.type === "paragraph" && borderParagraph.properties.borders?.bottom,
    ).toContain("currentColor");
  });

  it("splits page-break runs without duplicating list markers or empty segments", () => {
    const paragraph: Paragraph = {
      type: "paragraph",
      properties: { before: 6 },
      list: { numId: "1", level: 0, marker: "1." },
      runs: [
        { type: "text", text: "A" },
        { type: "break", page: true },
        { type: "text", text: "B" },
      ],
    };
    const [first, second] = splitAtPageBreaks(paragraph);
    expect(first.runs).toEqual([{ type: "text", text: "A" }]);
    expect(first.properties.pageBreakAfter).toBe(true);
    expect(second.runs).toEqual([{ type: "text", text: "B" }]);
    expect(second.properties.before).toBe(0);
    expect(second.list).toBeUndefined();
    const trailing = splitAtPageBreaks({
      ...paragraph,
      runs: [
        { type: "text", text: "A" },
        { type: "break", page: true },
      ],
    });
    expect(trailing).toHaveLength(1);
    expect(trailing[0].properties.pageBreakAfter).toBe(true);
    expect(trailing[0].runs).toEqual([{ type: "text", text: "A" }]);
  });
});
