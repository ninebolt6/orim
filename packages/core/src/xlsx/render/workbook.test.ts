import { describe, expect, it, vi } from "vite-plus/test";
import type { Worksheet, XlsxWorkbook } from "../model";
import { renderWorkbook, renderXlsx } from "./workbook";

const worksheet = (name: string, state: Worksheet["state"] = "visible"): Worksheet => ({
  name,
  state,
  rows: [],
  columns: [],
  merges: [],
  hyperlinks: [],
  defaultColumnWidth: 70,
  defaultRowHeight: 15,
  frozen: { rows: 0, columns: 0 },
  showGridLines: true,
  showHeaders: true,
});

const workbook = (sheets: readonly Worksheet[], activeSheet = 0): XlsxWorkbook => ({
  sheets,
  activeSheet,
  date1904: false,
});

describe("XLSX workbook rendering", () => {
  it("renders accessible sheet tabs and cancels a prior render on container reuse", () => {
    const container = document.createElement("div");
    const first = workbook([
      worksheet('<img src=x onerror="alert(1)">'),
      worksheet("Hidden", "hidden"),
      worksheet("Data"),
    ]);
    const onSheetChange = vi.fn();
    const rendered = renderWorkbook(first, container, { onSheetChange });

    expect(container.querySelectorAll('[role="tab"]')).toHaveLength(2);
    expect(container.textContent).toContain('<img src=x onerror="alert(1)">');
    expect(container.querySelector("img")).toBeNull();
    container.querySelector<HTMLButtonElement>('[role="tab"]')!.click();
    expect(rendered.sheet).toBe(0);
    expect(onSheetChange).toHaveBeenCalledWith(0);
    container.querySelector<HTMLButtonElement>('[role="tab"]:last-child')!.click();
    expect(rendered.sheet).toBe(2);
    expect(onSheetChange).toHaveBeenCalledWith(2);
    rendered.update({ sheet: 2 });
    rendered.update({});

    const replacement = renderWorkbook(workbook([worksheet("Replacement")]), container);
    expect(container.querySelectorAll(".orim-xlsx-book")).toHaveLength(1);
    rendered.destroy();
    rendered.update({ sheet: 0 });
    expect(container.querySelector(".orim-xlsx-book")).not.toBeNull();
    replacement.destroy();
    expect(container.childElementCount).toBe(0);
  });

  it("renders a no-visible-worksheets notice for a model without visible sheets", () => {
    const container = document.createElement("div");
    const rendered = renderWorkbook(
      workbook([worksheet("Hidden", "hidden"), worksheet("Very hidden", "veryHidden")]),
      container,
    );

    expect(container.textContent).toContain("No visible worksheets.");
    rendered.destroy();
  });

  it("rejects hidden and out-of-range sheet selections", () => {
    const container = document.createElement("div");
    const model = workbook([worksheet("Visible"), worksheet("Hidden", "hidden")]);
    expect(() => renderWorkbook(model, container, { sheet: 1 })).toThrow(RangeError);
    expect(() => renderWorkbook(model, container, { sheet: 2 })).toThrow(RangeError);
    expect(() => renderWorkbook(model, container, { sheet: -1 })).toThrow(RangeError);
    expect(() => renderWorkbook(model, container, { sheet: 0.5 })).toThrow(RangeError);
    expect(() => renderWorkbook(model, container, { sheet: Number.NaN })).toThrow(RangeError);
  });

  it("rejects an already-aborted request and cancels parsing after abort", async () => {
    const container = document.createElement("div");
    const alreadyAborted = new AbortController();
    alreadyAborted.abort();
    await expect(
      renderXlsx(new Uint8Array(), container, { signal: alreadyAborted.signal }),
    ).rejects.toMatchObject({ name: "AbortError" });

    const controller = new AbortController();
    const pending = renderXlsx(new Uint8Array(), container, { signal: controller.signal });
    controller.abort();
    await expect(pending).rejects.toMatchObject({ name: "AbortError" });
  });

  it("surfaces package errors from asynchronous parsing without leaving partial UI", async () => {
    const container = document.createElement("div");

    await expect(renderXlsx(new Uint8Array(), container)).rejects.toThrow();
    expect(container.childElementCount).toBe(0);
  });
});
