import { afterEach, describe, expect, it, vi } from "vite-plus/test";
import { StrictMode } from "react";
import { cleanup, fireEvent, render } from "@testing-library/react";
import * as core from "@orim/core";
import { XlsxViewer } from "./XlsxViewer";
import type { XlsxWorkbook } from "@orim/core";

const workbook: XlsxWorkbook = {
  activeSheet: 0,
  date1904: false,
  sheets: [
    {
      name: "Summary",
      state: "visible",
      rows: [],
      columns: [],
      merges: [],
      hyperlinks: [],
      defaultColumnWidth: 70,
      defaultRowHeight: 15,
      frozen: { rows: 0, columns: 0 },
      showGridLines: true,
      showHeaders: true,
    },
    {
      name: "Hidden",
      state: "hidden",
      rows: [],
      columns: [],
      merges: [],
      hyperlinks: [],
      defaultColumnWidth: 70,
      defaultRowHeight: 15,
      frozen: { rows: 0, columns: 0 },
      showGridLines: true,
      showHeaders: true,
    },
    {
      name: "Data",
      state: "visible",
      rows: [],
      columns: [],
      merges: [],
      hyperlinks: [],
      defaultColumnWidth: 70,
      defaultRowHeight: 15,
      frozen: { rows: 0, columns: 0 },
      showGridLines: true,
      showHeaders: true,
    },
  ],
};

describe("XlsxViewer", () => {
  afterEach(() => {
    cleanup();
    vi.restoreAllMocks();
  });

  it("renders only visible tabs and switches the rendered worksheet", () => {
    const onSheetChange = vi.fn();
    const view = render(<XlsxViewer workbook={workbook} onSheetChange={onSheetChange} />);
    expect(view.getAllByRole("tab").map((tab) => tab.textContent)).toEqual(["Summary", "Data"]);
    expect(view.container.querySelector(".orim-xlsx-grid")).not.toBeNull();
    fireEvent.click(view.getByRole("tab", { name: "Data" }));
    expect(onSheetChange).toHaveBeenCalledWith(2);
    expect(view.getByRole("tab", { name: "Data" }).getAttribute("aria-selected")).toBe("true");
  });

  it("uses controlled selection and destroys the previous worksheet", () => {
    const spy = vi.spyOn(core, "renderWorksheet");
    const view = render(<XlsxViewer workbook={workbook} sheet={0} />);
    const first = spy.mock.results[0]!.value;
    const destroy = vi.spyOn(first, "destroy");
    view.rerender(<XlsxViewer workbook={workbook} sheet={2} />);
    expect(destroy).toHaveBeenCalledOnce();
    expect(spy.mock.calls.map(([, sheet]) => sheet)).toEqual([0, 2]);
    expect(view.getByRole("tab", { name: "Data" }).getAttribute("aria-selected")).toBe("true");
  });

  it("does not render a missing workbook and cleans up on unmount", () => {
    const spy = vi.spyOn(core, "renderWorksheet");
    const view = render(<XlsxViewer workbook={null} />);
    expect(spy).not.toHaveBeenCalled();
    view.rerender(<XlsxViewer workbook={workbook} />);
    const rendered = spy.mock.results[0]!.value;
    const destroy = vi.spyOn(rendered, "destroy");
    view.unmount();
    expect(destroy).toHaveBeenCalledOnce();
  });

  it("resets uncontrolled selection to a replacement workbook's active sheet", () => {
    const view = render(<XlsxViewer workbook={workbook} />);
    fireEvent.click(view.getByRole("tab", { name: "Data" }));
    const replacement = { ...workbook, activeSheet: 0 };
    view.rerender(<XlsxViewer workbook={replacement} />);
    expect(view.getByRole("tab", { name: "Summary" }).getAttribute("aria-selected")).toBe("true");
  });

  it("keeps one grid under React StrictMode", () => {
    const view = render(
      <StrictMode>
        <XlsxViewer workbook={workbook} />
      </StrictMode>,
    );
    expect(view.container.querySelectorAll(".orim-xlsx-grid")).toHaveLength(1);
  });
});
