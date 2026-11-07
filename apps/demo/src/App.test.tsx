import { afterEach, describe, expect, it, vi } from "vite-plus/test";
import { StrictMode } from "react";
import { act, cleanup, fireEvent, render } from "@testing-library/react";
import { parseDocx, parseXlsx } from "@orim/core";
import { App } from "./App";
import type { LoadResult } from "./load";

vi.mock("@orim/core", () => ({ parseDocx: vi.fn(), parseXlsx: vi.fn() }));
vi.mock("@orim/react", () => ({
  DocxViewer: () => <div>Word document</div>,
  XlsxViewer: () => <div>Excel workbook</div>,
}));

const deferredLoad = () => {
  let resolve!: (result: LoadResult) => void;
  const promise = new Promise<LoadResult>((complete) => {
    resolve = complete;
  });
  return { promise, resolve };
};

describe("demo file loading", () => {
  afterEach(() => {
    cleanup();
    vi.resetAllMocks();
  });

  it("shows the toolbar and loading state immediately, then the initial sample in StrictMode", async () => {
    const initial = deferredLoad();
    const view = render(
      <StrictMode>
        <App initialLoad={initial.promise} />
      </StrictMode>,
    );
    expect(view.getByText("Open file")).not.toBeNull();
    expect(view.getByRole("status").textContent).toBe("Loading file…");
    await act(async () => {
      initial.resolve({
        view: { kind: "xlsx", workbook: { sheets: [], activeSheet: 0, date1904: false } },
        error: "",
      });
    });
    expect(view.queryByRole("status")).toBeNull();
    expect(view.getByText("Excel workbook")).not.toBeNull();
  });

  it("shows startup errors and an empty initial result without hanging", async () => {
    const view = render(<App initialLoad={Promise.resolve({ view: null, error: "Offline" })} />);
    await act(async () => {});
    expect(view.getByRole("alert").textContent).toBe("Offline");
    view.rerender(<App initialLoad={Promise.resolve({ view: null, error: "" })} />);
    await act(async () => {});
    expect(view.getByText("Open a DOCX or XLSX file.")).not.toBeNull();
  });

  it("rejects unsupported dropped files and does not let startup overwrite the error", async () => {
    const initial = deferredLoad();
    const view = render(<App initialLoad={initial.promise} />);
    fireEvent.drop(view.container.firstElementChild!, {
      dataTransfer: { files: [new File([], "old.xls")] },
    });
    await act(async () => {});
    expect(view.getByRole("alert").textContent).toContain("Open a .docx or .xlsx file.");
    expect(parseDocx).not.toHaveBeenCalled();
    expect(parseXlsx).not.toHaveBeenCalled();
    await act(async () => {
      initial.resolve({ view: null, error: "Startup error" });
    });
    expect(view.getByRole("alert").textContent).not.toContain("Startup error");
    expect(view.queryByRole("status")).toBeNull();
  });

  it("accepts uppercase supported extensions and displays parser errors", async () => {
    const view = render(<App initialLoad={Promise.resolve({ view: null, error: "" })} />);
    const input = view.container.querySelector("input[type=file]")!;
    expect(input.getAttribute("accept")).toBe(".docx,.xlsx");
    fireEvent.change(input, { target: { files: [new File([], "book.XLSX")] } });
    await act(async () => {});
    expect(parseXlsx).toHaveBeenCalledOnce();
    expect(view.getByText("Excel workbook")).not.toBeNull();
    vi.mocked(parseDocx).mockRejectedValue(new Error("Password protected"));
    fireEvent.change(input, { target: { files: [new File([], "word.DOCX")] } });
    await act(async () => {});
    expect(parseDocx).toHaveBeenCalledOnce();
    expect(view.getByRole("alert").textContent).toBe("Password protected");
  });
});
