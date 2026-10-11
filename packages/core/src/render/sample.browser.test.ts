import { describe, expect, it } from "vite-plus/test";
import { parseDocx } from "../docx/parse";
import type { DocxDocument } from "../docx/model";
import { renderDocument } from "../index";
import "./styles.css";
import sampleUrl from "../../../../fixtures/sample.docx?url";

const sampleDocument = async (): Promise<DocxDocument> => {
  const response = await fetch(sampleUrl);
  return parseDocx(await response.blob());
};

const createContainer = (): HTMLDivElement => {
  const container = document.createElement("div");
  document.body.append(container);
  return container;
};

const deletedTextDocument = (): DocxDocument => {
  return {
    sections: [
      {
        properties: {
          width: 595,
          height: 842,
          orientation: "portrait",
          margins: { top: 48, right: 48, bottom: 48, left: 48, header: 20, footer: 20 },
        },
        blocks: Array.from({ length: 80 }, () => ({
          type: "paragraph" as const,
          properties: {},
          runs: [
            {
              type: "text" as const,
              text: "Deleted content fills this paragraph with enough repeated words to wrap across several lines. ".repeat(
                5,
              ),
              change: { kind: "del" as const, author: "Aiko", date: "" },
            },
          ],
        })),
      },
    ],
    comments: [],
    imageParts: {},
  };
};

const tableAfterContentDocument = (): DocxDocument => {
  const paragraph = (text: string) => ({
    type: "paragraph" as const,
    properties: {},
    runs: [{ type: "text" as const, text }],
  });
  return {
    sections: [
      {
        properties: {
          width: 595,
          height: 842,
          orientation: "portrait",
          margins: { top: 48, right: 48, bottom: 48, left: 48, header: 20, footer: 20 },
        },
        blocks: [
          ...Array.from({ length: 24 }, (_, index) =>
            paragraph(`Leading paragraph ${index + 1} `.repeat(8)),
          ),
          {
            type: "table",
            widths: ["50%", "50%"],
            rows: [
              {
                header: true,
                cells: [
                  { gridSpan: 1, rowSpan: 1, blocks: [paragraph("Header A")] },
                  { gridSpan: 1, rowSpan: 1, blocks: [paragraph("Header B")] },
                ],
              },
              ...Array.from({ length: 30 }, (_, index) => ({
                cells: [
                  { gridSpan: 1, rowSpan: 1, blocks: [paragraph(`Body row ${index + 1}`)] },
                  { gridSpan: 1, rowSpan: 1, blocks: [paragraph(`Value ${index + 1}`)] },
                ],
              })),
            ],
          },
        ],
      },
    ],
    comments: [],
    imageParts: {},
  };
};

describe("sample DOCX in Chromium", () => {
  it("uses portrait and landscape page geometry without overflow", async () => {
    const container = createContainer();
    const rendered = renderDocument(await sampleDocument(), container);
    const pages = Array.from(container.querySelectorAll<HTMLElement>(".orim-page"));
    expect(pages).toHaveLength(7);
    expect(pages[0].offsetWidth).toBeLessThan(pages[0].offsetHeight);
    expect(pages.at(-1)!.offsetWidth).toBeGreaterThan(pages.at(-1)!.offsetHeight);
    expect(Math.abs(pages[0].offsetWidth - (595.3 * 96) / 72)).toBeLessThan(2);
    for (const page of pages) {
      const content = page.querySelector<HTMLElement>(".orim-content")!;
      expect(content.scrollHeight).toBeLessThanOrEqual(content.clientHeight + 1);
    }
    rendered.destroy();
    container.remove();
  });

  it("continues the long table and repeats its header", async () => {
    const container = createContainer();
    const rendered = renderDocument(await sampleDocument(), container);
    const tableParts = Array.from(container.querySelectorAll<HTMLTableElement>(".orim-table"));
    expect(tableParts.length).toBeGreaterThanOrEqual(2);
    expect(
      tableParts.filter((table) => table.querySelector('[data-orim-header="true"]')).length,
    ).toBeGreaterThanOrEqual(2);
    rendered.destroy();
    container.remove();
  });

  it("conserves table body rows when a table follows preceding content", async () => {
    const container = createContainer();
    const rendered = renderDocument(tableAfterContentDocument(), container);
    const tables = Array.from(container.querySelectorAll<HTMLTableElement>(".orim-table"));
    const rows = tables.flatMap((table) =>
      Array.from(table.rows)
        .filter((row) => row.dataset.orimHeader !== "true")
        .map((row) => row.textContent?.trim() ?? ""),
    );
    expect(tables.length).toBeGreaterThan(1);
    expect(
      tables.filter((table) => table.querySelector('[data-orim-header="true"]')).length,
    ).toBeGreaterThan(1);
    expect(rows).toEqual(
      Array.from({ length: 30 }, (_, index) => `Body row ${index + 1}Value ${index + 1}`),
    );
    expect(new Set(rows).size).toBe(30);
    rendered.destroy();
    container.remove();
  });

  it("keeps the bullet marker visible inside the page", async () => {
    const container = createContainer();
    const rendered = renderDocument(await sampleDocument(), container);
    const paragraph = Array.from(container.querySelectorAll<HTMLElement>(".orim-paragraph")).find(
      (item) => item.textContent?.includes("Bullet item"),
    )!;
    const list = rendered.document.sections
      .flatMap((section) => section.blocks)
      .find(
        (block) =>
          block.type === "paragraph" &&
          block.runs.some((run) => run.type === "text" && run.text.includes("Bullet item")),
      );
    if (list?.type !== "paragraph" || !list.list?.left || list.list.hanging === undefined)
      throw new Error("The fixture bullet needs explicit left and hanging indents.");
    const marker = paragraph.querySelector<HTMLElement>(".orim-list-marker")!;
    const page = paragraph.closest<HTMLElement>(".orim-page")!;
    const content = paragraph.closest<HTMLElement>(".orim-content")!;
    const markerRect = marker.getBoundingClientRect();
    const pageRect = page.getBoundingClientRect();
    const contentRect = content.getBoundingClientRect();
    const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
    let bulletText: Text | undefined;
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      if (node.textContent?.includes("Bullet item")) {
        bulletText = node as Text;
        break;
      }
    }
    expect(bulletText).toBeDefined();
    const range = document.createRange();
    range.setStart(bulletText!, bulletText!.textContent!.indexOf("Bullet item"));
    range.setEnd(bulletText!, bulletText!.textContent!.indexOf("Bullet item") + 1);
    const textRect = range.getBoundingClientRect();
    expect(markerRect.width).toBeGreaterThan(0);
    expect(markerRect.height).toBeGreaterThan(0);
    expect(markerRect.right).toBeGreaterThan(pageRect.left);
    expect(markerRect.left).toBeLessThan(pageRect.right);
    expect(markerRect.bottom).toBeGreaterThan(pageRect.top);
    expect(markerRect.top).toBeLessThan(pageRect.bottom);
    expect(markerRect.left).toBeGreaterThanOrEqual(contentRect.left - 0.5);
    expect(
      Math.abs(
        markerRect.left - (contentRect.left + ((list.list.left - list.list.hanging) * 96) / 72),
      ),
    ).toBeLessThanOrEqual(1);
    expect(
      Math.abs(textRect.left - (contentRect.left + (list.list.left * 96) / 72)),
    ).toBeLessThanOrEqual(1);
    rendered.destroy();
    container.remove();
  });

  it("keeps comment balloons near their anchors inside their page", async () => {
    const container = createContainer();
    const rendered = renderDocument(await sampleDocument(), container);
    const balloons = Array.from(
      container.querySelectorAll<HTMLElement>(".orim-page-comments > .orim-comment"),
    );
    expect(balloons.length).toBeGreaterThan(0);
    const previousByPage = new Map<HTMLElement, DOMRect>();
    for (const balloon of balloons) {
      const wrapper = balloon.closest<HTMLElement>(".orim-page-wrap")!;
      const page = wrapper.querySelector<HTMLElement>(".orim-page")!;
      const commentId = balloon.dataset.comment;
      const anchor = Array.from(
        page.querySelectorAll<HTMLElement>(".orim-comment-anchor[data-c]"),
      ).find((item) => item.dataset.c?.split(" ").includes(commentId ?? ""))!;
      const balloonRect = balloon.getBoundingClientRect();
      const anchorRect = anchor.getBoundingClientRect();
      const wrapperRect = wrapper.getBoundingClientRect();
      expect(balloonRect.top).toBeGreaterThanOrEqual(anchorRect.top - 2);
      expect(balloonRect.top).toBeGreaterThanOrEqual(wrapperRect.top - 1);
      expect(balloonRect.bottom).toBeLessThanOrEqual(wrapperRect.bottom + 1);
      const previous = previousByPage.get(wrapper);
      if (!previous || anchorRect.top >= previous.bottom + 8)
        expect(Math.abs(balloonRect.top - anchorRect.top)).toBeLessThanOrEqual(4);
      previousByPage.set(wrapper, balloonRect);
    }
    rendered.destroy();
    container.remove();
  });

  it("starts the paragraph after an explicit page break at the next page top", async () => {
    const container = createContainer();
    const rendered = renderDocument(await sampleDocument(), container);
    const paragraphs = Array.from(container.querySelectorAll<HTMLElement>(".orim-paragraph"));
    const before = paragraphs.find((paragraph) =>
      paragraph.textContent?.includes("Before inline page break"),
    )!;
    const after = paragraphs.find((paragraph) =>
      paragraph.textContent?.includes("After inline page break"),
    )!;
    expect(before.closest(".orim-page-wrap")).not.toBe(after.closest(".orim-page-wrap"));
    expect(after.closest(".orim-content")?.firstElementChild).toBe(after);
    rendered.destroy();
    container.remove();
  });

  it("repaginates deleted text when switching to Final mode", async () => {
    const container = createContainer();
    const rendered = renderDocument(deletedTextDocument(), container);
    const markupPages = container.querySelectorAll(".orim-page").length;
    rendered.update({ changes: "final" });
    const finalPages = container.querySelectorAll(".orim-page").length;
    expect(markupPages).toBeGreaterThan(1);
    expect(finalPages).toBeLessThan(markupPages);
    rendered.destroy();
    container.remove();
  });
});
