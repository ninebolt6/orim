import { afterEach, describe, expect, it, vi } from "vite-plus/test";
import { parseDocx, parsePptx, parseXlsx } from "@orim/core";
import { loadSample, parseFile } from "./load";

vi.mock("@orim/core", () => ({ parseDocx: vi.fn(), parsePptx: vi.fn(), parseXlsx: vi.fn() }));

describe("demo loaders", () => {
  afterEach(() => {
    vi.resetAllMocks();
    vi.unstubAllGlobals();
  });

  it("dispatches supported extensions to their respective parsers", async () => {
    const file = new Blob();
    await expect(parseFile(file, "docx")).resolves.toMatchObject({ kind: "docx" });
    expect(parseDocx).toHaveBeenCalledWith(file);
    expect(parseXlsx).not.toHaveBeenCalled();
    await expect(parseFile(file, "xlsx")).resolves.toMatchObject({ kind: "xlsx" });
    expect(parseXlsx).toHaveBeenCalledWith(file);
    await expect(parseFile(file, "pptx")).resolves.toMatchObject({ kind: "pptx" });
    expect(parsePptx).toHaveBeenCalledWith(file);
  });

  it("rejects unsupported and missing extensions before parsing", async () => {
    for (const extension of ["doc", "xls", "ppt", "pdf", ""]) {
      await expect(parseFile(new Blob(), extension)).rejects.toThrow(
        "Open a .docx, .xlsx or .pptx file.",
      );
    }
    expect(parseDocx).not.toHaveBeenCalled();
    expect(parseXlsx).not.toHaveBeenCalled();
    expect(parsePptx).not.toHaveBeenCalled();
  });

  it("loads samples and propagates fetch and parsing failures", async () => {
    const blob = new Blob();
    const fetcher = vi.fn().mockResolvedValue({ ok: true, blob: async () => blob });
    vi.stubGlobal("fetch", fetcher);
    await expect(loadSample("xlsx")).resolves.toMatchObject({ kind: "xlsx" });
    expect(parseXlsx).toHaveBeenCalledWith(blob);
    await expect(loadSample("pptx")).resolves.toMatchObject({ kind: "pptx" });
    expect(parsePptx).toHaveBeenCalledWith(blob);
    fetcher.mockResolvedValue({ ok: false, status: 404 });
    await expect(loadSample("docx")).rejects.toThrow("Unable to load sample: 404");
    fetcher.mockRejectedValue(new Error("Offline"));
    await expect(loadSample("docx")).rejects.toThrow("Offline");
    fetcher.mockResolvedValue({ ok: true, blob: async () => blob });
    vi.mocked(parseDocx).mockRejectedValue(new Error("Invalid document"));
    await expect(loadSample("docx")).rejects.toThrow("Invalid document");
  });
});
