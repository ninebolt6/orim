import { describe, expect, it } from "vite-plus/test";
import { parseDocx } from "../src/docx/parse";
import { parsePptx } from "../src/pptx/parse";
import { parseXlsx } from "../src/xlsx/parse";
import pythonDocxUrl from "./fixtures/python-docx-basic.docx?url";
import npmDocxUrl from "./fixtures/npm-docx.docx?url";
import pythonPptxUrl from "./fixtures/python-pptx-basic.pptx?url";
import pptxgenjsUrl from "./fixtures/pptxgenjs.pptx?url";
import sheetjsUrl from "./fixtures/sheetjs.xlsx?url";
import xlsxwriterUrl from "./fixtures/xlsxwriter.xlsx?url";
import strictUrl from "./fixtures/strict.xlsx?url";

const blob = async (url: string): Promise<Blob> => (await fetch(url)).blob();

describe("files from third-party generators", () => {
  it("parses a python-docx document with tables, lists and styled runs", async () => {
    const doc = await parseDocx(await blob(pythonDocxUrl));
    const text = (block: (typeof doc.sections)[number]["blocks"][number]): string =>
      block.type === "paragraph"
        ? block.runs.map((run) => (run.type === "text" ? run.text : "")).join("")
        : "";
    const blocks = doc.sections[0]!.blocks;
    expect(blocks[0]).toMatchObject({ type: "paragraph", properties: { styleId: "Heading1" } });
    expect(text(blocks[0]!)).toBe("Heading 1");
    const styled = blocks.find((block) => text(block).includes("colored"));
    expect(
      styled?.type === "paragraph" &&
        styled.runs.some((run) => run.type === "text" && run.style?.bold === true),
    ).toBe(true);
  });

  it("parses an npm-docx document with hyperlinks, lists and a table", async () => {
    const doc = await parseDocx(await blob(npmDocxUrl));
    const blocks = doc.sections[0]!.blocks;
    expect(blocks.some((block) => block.type === "table")).toBe(true);
    const hyperlink = blocks
      .flatMap((block) => (block.type === "paragraph" ? block.runs : []))
      .find((run) => run.type === "text" && run.href === "https://example.com");
    expect(hyperlink).toBeTruthy();
    expect(blocks.some((block) => block.type === "paragraph" && block.list !== undefined)).toBe(
      true,
    );
  });

  it("parses a python-pptx presentation with slide layouts", async () => {
    const pres = await parsePptx(await blob(pythonPptxUrl));
    expect(pres.slides).toHaveLength(4);
    expect(pres.width).toBeGreaterThan(0);
    expect(pres.slides[0]!.elements.length).toBeGreaterThan(0);
    const texts = pres.slides.flatMap((slide) =>
      slide.elements.flatMap((el) =>
        el.kind === "shape" && el.text
          ? el.text.paragraphs.flatMap((p) => p.runs.map((r) => ("text" in r ? r.text : "")))
          : [],
      ),
    );
    expect(texts.some((t) => t.includes("Layout 0"))).toBe(true);
  });

  it("parses a pptxgenjs presentation with shapes, table and image", async () => {
    const pres = await parsePptx(await blob(pptxgenjsUrl));
    expect(pres.slides.length).toBeGreaterThanOrEqual(3);
    const kinds = new Set(pres.slides.flatMap((slide) => slide.elements.map((el) => el.kind)));
    expect(kinds.has("table")).toBe(true);
    expect(kinds.has("picture")).toBe(true);
  });

  it("parses a SheetJS workbook with formulas and a hyperlink", async () => {
    const wb = await parseXlsx(await blob(sheetjsUrl));
    expect(wb.sheets.map((s) => s.name)).toEqual(["Sheet1", "Second"]);
    const sheet = wb.sheets[0]!;
    expect(sheet.rows[0]!.cells[0]!.value).toEqual({ type: "string", value: "Name" });
    expect(sheet.rows[1]!.cells[1]!.value).toEqual({ type: "number", value: 3 });
  });

  it("parses an XlsxWriter workbook with a hyperlink and merged cells", async () => {
    const wb = await parseXlsx(await blob(xlsxwriterUrl));
    const sheet = wb.sheets.find((s) => s.name === "Data")!;
    expect(sheet.hyperlinks).toEqual([
      { range: { top: 5, left: 0, bottom: 5, right: 0 }, href: "https://example.com" },
    ]);
    expect(sheet.merges).toContainEqual({ top: 0, left: 4, bottom: 0, right: 6 });
    expect(sheet.frozen).toEqual({ rows: 1, columns: 1 });
  });

  it("parses a strict-namespace workbook", async () => {
    const wb = await parseXlsx(await blob(strictUrl));
    expect(wb.sheets).toHaveLength(1);
    expect(wb.sheets[0]!.rows[0]!.cells[0]!.value).toEqual({ type: "string", value: "text" });
  });
});
