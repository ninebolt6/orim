import { describe, expect, it } from "vite-plus/test";
import { parseDocx, renderDocx } from "../index";
import sampleUrl from "../../../../fixtures/sample.docx?url";

describe("sample DOCX in Chromium", () => {
  it("renders seven pages with portrait and landscape geometry", async () => {
    const response = await fetch(sampleUrl);
    const doc = await parseDocx(await response.blob());
    const container = document.createElement("div");
    document.body.append(container);
    const rendered = await renderDocx(doc, container);
    const pages = Array.from(container.querySelectorAll<HTMLElement>(".orim-page"));
    expect(pages).toHaveLength(7);
    expect(pages[0].offsetWidth).toBeLessThan(pages[0].offsetHeight);
    expect(pages.at(-1)!.offsetWidth).toBeGreaterThan(pages.at(-1)!.offsetHeight);
    expect(Math.abs(pages[0].offsetWidth - (595.3 * 96) / 72)).toBeLessThan(2);
    for (const page of pages) {
      const content = page.querySelector<HTMLElement>(".orim-content")!;
      expect(content.scrollHeight).toBeLessThanOrEqual(content.clientHeight + 1);
    }
    const tableParts = Array.from(container.querySelectorAll<HTMLTableElement>(".orim-table"));
    expect(tableParts.length).toBeGreaterThanOrEqual(2);
    expect(
      tableParts.filter((table) => table.querySelector('[data-orim-header="true"]')).length,
    ).toBeGreaterThanOrEqual(2);
    expect(container.querySelectorAll(".orim-comment").length).toBeGreaterThan(0);
    rendered.destroy();
    container.remove();
  });
});
