import { describe, expect, it } from "vite-plus/test";
import { strToU8, zipSync } from "fflate";
import { openZip } from "./opc/zip";

describe("ZIP reader", () => {
  it("reads stored and deflated files", async () => {
    const stored = zipSync({ "stored.txt": strToU8("stored") }, { level: 0 });
    const deflated = zipSync({ "deflated.txt": strToU8("deflated content") }, { level: 6 });
    const archive = openZip(new Uint8Array(stored));
    expect(await archive.text("stored.txt")).toBe("stored");
    expect(await openZip(new Uint8Array(deflated)).text("deflated.txt")).toBe("deflated content");
  });

  it("reports a missing file clearly", async () => {
    const archive = openZip(new Uint8Array(zipSync({ "present.txt": strToU8("ok") })));
    await expect(archive.read("missing.txt")).rejects.toThrow("ZIP entry not found: missing.txt");
  });
});
