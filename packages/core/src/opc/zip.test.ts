import { describe, expect, it } from "vite-plus/test";
import { strToU8, zipSync } from "fflate";
import { openZip } from "./zip";

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

  it("rejects unsupported compression methods", async () => {
    const bytes = new Uint8Array(zipSync({ "file.txt": strToU8("content") }, { level: 0 }));
    const central = findSignature(bytes, 0x02014b50);
    new DataView(bytes.buffer).setUint16(central + 10, 12, true);
    await expect(openZip(bytes).read("file.txt")).rejects.toThrow(
      "Unsupported ZIP compression method 12: file.txt",
    );
  });

  it("rejects deflated output larger than its declared size", async () => {
    const bytes = new Uint8Array(
      zipSync({ "file.txt": strToU8("expanded content") }, { level: 6 }),
    );
    const central = findSignature(bytes, 0x02014b50);
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    view.setUint32(central + 24, 1, true);
    await expect(openZip(bytes).read("file.txt")).rejects.toThrow(
      "ZIP entry exceeds its declared uncompressed size.",
    );
  });

  it("rejects encrypted entries", () => {
    const bytes = new Uint8Array(zipSync({ "file.txt": strToU8("content") }, { level: 0 }));
    const central = findSignature(bytes, 0x02014b50);
    new DataView(bytes.buffer).setUint16(
      central + 8,
      new DataView(bytes.buffer).getUint16(central + 8, true) | 1,
      true,
    );
    expect(() => openZip(bytes)).toThrow("Encrypted ZIP entries are not supported.");
  });

  it("rejects ZIP64 archives explicitly", () => {
    const bytes = new Uint8Array(zipSync({ "file.txt": strToU8("content") }, { level: 0 }));
    const eocd = findSignature(bytes, 0x06054b50);
    new DataView(bytes.buffer).setUint16(eocd + 10, 0xffff, true);
    expect(() => openZip(bytes)).toThrow("ZIP64 and multi-disk archives are not supported.");
  });
});

function findSignature(bytes: Uint8Array, signature: number): number {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  for (let offset = 0; offset <= bytes.byteLength - 4; offset++)
    if (view.getUint32(offset, true) === signature) return offset;
  throw new Error("Signature not found.");
}
