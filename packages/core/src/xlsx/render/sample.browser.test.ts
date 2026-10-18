import { describe, expect, it } from "vite-plus/test";
import { parseXlsx } from "../parse";
import { computeLayout } from "../layout";
import { renderWorkbook } from "./workbook";
import "./styles.css";
import sampleUrl from "../../../../../fixtures/sample.xlsx?url";

const sampleWorkbook = async () => {
  const response = await fetch(sampleUrl);
  return parseXlsx(await response.blob());
};

describe("sample XLSX in Chromium", () => {
  it("preserves widths, frozen headers, spill, row heights, and hidden dimensions", async () => {
    const workbook = await sampleWorkbook();
    const container = document.createElement("div");
    container.style.height = "220px";
    container.style.width = "520px";
    document.body.append(container);
    const rendered = renderWorkbook(workbook, container);
    const scroll = container.querySelector<HTMLElement>(".orim-xlsx-scroll")!;
    const summary = workbook.sheets[0]!;
    const columns = Array.from(container.querySelectorAll<HTMLElement>("colgroup col")).slice(1);
    expect(columns.map((column) => Number.parseFloat(column.style.width))).toEqual(
      computeLayout(summary).columns.map((column) => column.width),
    );
    expect(columns).toHaveLength(6);
    const hiddenColumn = Array.from(
      container.querySelectorAll<HTMLElement>(".orim-xlsx-column-header"),
    ).some((header) => header.textContent === "G");
    expect(hiddenColumn).toBe(false);
    expect(
      Array.from(container.querySelectorAll<HTMLElement>(".orim-xlsx-row-header")).some(
        (header) => header.textContent === "10",
      ),
    ).toBe(false);
    expect(
      Array.from(container.querySelectorAll('[role="tab"]')).map((tab) => tab.textContent),
    ).toEqual(["Summary", "Data", "No grid"]);
    const spill = container.querySelector<HTMLElement>(".orim-xlsx-spill")!;
    const spillContent = spill.querySelector<HTMLElement>(".orim-xlsx-content")!;
    expect(spillContent.scrollWidth).toBeGreaterThan(spill.clientWidth);
    const wrappedRow = Array.from(container.querySelectorAll<HTMLTableRowElement>("tbody tr")).find(
      (row) => row.textContent?.includes("This long wrapped cell stays within"),
    )!;
    const wrappedContent = wrappedRow.querySelector<HTMLElement>(".orim-xlsx-content")!;
    expect(wrappedRow.style.height).toBe("45pt");
    expect(wrappedContent.getBoundingClientRect().height).toBeCloseTo(60, 0);
    scroll.scrollTop = 100;
    const frozenHeader = container.querySelector<HTMLElement>(".orim-xlsx-column-header")!;
    expect(frozenHeader.getBoundingClientRect().top).toBeGreaterThanOrEqual(
      scroll.getBoundingClientRect().top,
    );
    rendered.destroy();
    container.remove();
  });
});
