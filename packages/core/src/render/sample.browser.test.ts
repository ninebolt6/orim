import { describe, expect, it } from "vite-plus/test";
import { parseDocx } from "../docx/parse";
import type { DocxDocument } from "../docx/model";
import { renderDocx } from "../index";
import sampleUrl from "../../../../fixtures/sample.docx?url";

async function sampleDocument(): Promise<DocxDocument> {
  const response = await fetch(sampleUrl);
  return parseDocx(await response.blob());
}

function createContainer(): HTMLDivElement {
  const container = document.createElement("div");
  document.body.append(container);
  return container;
}

function deletedTextDocument(): DocxDocument {
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
}

describe("sample DOCX in Chromium", () => {
  it("uses portrait and landscape page geometry without overflow", async () => {
    const container = createContainer();
    const rendered = await renderDocx(await sampleDocument(), container);
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
    const rendered = await renderDocx(await sampleDocument(), container);
    const tableParts = Array.from(container.querySelectorAll<HTMLTableElement>(".orim-table"));
    expect(tableParts.length).toBeGreaterThanOrEqual(2);
    expect(
      tableParts.filter((table) => table.querySelector('[data-orim-header="true"]')).length,
    ).toBeGreaterThanOrEqual(2);
    rendered.destroy();
    container.remove();
  });

  it("keeps comment balloons near their anchors inside their page", async () => {
    const container = createContainer();
    const rendered = await renderDocx(await sampleDocument(), container);
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
    const rendered = await renderDocx(await sampleDocument(), container);
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
    const rendered = await renderDocx(deletedTextDocument(), container);
    const markupPages = container.querySelectorAll(".orim-page").length;
    rendered.update({ changes: "final" });
    const finalPages = container.querySelectorAll(".orim-page").length;
    expect(markupPages).toBeGreaterThan(1);
    expect(finalPages).toBeLessThan(markupPages);
    rendered.destroy();
    container.remove();
  });
});
