import { afterEach, describe, expect, it, vi } from "vite-plus/test";
import { parseDocx, parseXlsx } from "@orim/core";
import { loadSample, parseFile } from "./load";

vi.mock("@orim/core", () => ({ parseDocx: vi.fn(), parseXlsx: vi.fn() }));

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
  });

  it("rejects unsupported and missing extensions before parsing", async () => {
    for (const extension of ["doc", "xls", "pdf", ""]) {
      await expect(parseFile(new Blob(), extension)).rejects.toThrow("Open a .docx or .xlsx file.");
    }
    expect(parseDocx).not.toHaveBeenCalled();
    expect(parseXlsx).not.toHaveBeenCalled();
  });

  it("loads samples and propagates fetch and parsing failures", async () => {
    const blob = new Blob();
    const fetcher = vi.fn().mockResolvedValue({ ok: true, blob: async () => blob });
    vi.stubGlobal("fetch", fetcher);
    await expect(loadSample("xlsx")).resolves.toMatchObject({ kind: "xlsx" });
    expect(parseXlsx).toHaveBeenCalledWith(blob);
    fetcher.mockResolvedValue({ ok: false, status: 404 });
    await expect(loadSample("docx")).rejects.toThrow("Unable to load sample: 404");
    fetcher.mockRejectedValue(new Error("Offline"));
    await expect(loadSample("docx")).rejects.toThrow("Offline");
    fetcher.mockResolvedValue({ ok: true, blob: async () => blob });
    vi.mocked(parseDocx).mockRejectedValue(new Error("Invalid document"));
    await expect(loadSample("docx")).rejects.toThrow("Invalid document");
  });
});
