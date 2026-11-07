import { describe, expect, it } from "vite-plus/test";
import { strToU8, zipSync } from "fflate";
import { openZip } from "./zip";

describe("ZIP reader", () => {
  it("explains legacy or encrypted Office containers for byte views and buffers", () => {
    const signature = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
    const padded = new Uint8Array([0, ...signature, 0]);
    for (const input of [signature.buffer, padded.subarray(1, 9)]) {
      expect(() => openZip(input)).toThrow(
        "Unsupported file: this is a legacy binary Office file (.doc/.xls/.ppt) or a password-protected document. Save it as .docx/.xlsx/.pptx without a password.",
      );
    }
    for (const input of [new Uint8Array(), signature.subarray(0, 7)]) {
      expect(() => openZip(input)).toThrow(
        "Invalid ZIP: end-of-central-directory record not found.",
      );
    }
  });

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

  it("bounds cumulative decompression across reads", async () => {
    const content = new Uint8Array(64 * 1024 * 1024);
    const bytes = zipSync(
      { a: content, b: content, c: content, d: content, e: new Uint8Array([1]) },
      { level: 1 },
    );
    const archive = openZip(bytes);
    for (const path of ["a", "b", "c", "d"])
      expect((await archive.read(path)).length).toBe(content.length);
    await expect(archive.read("e")).rejects.toThrow(
      "ZIP entry exceeds the decompressed size limit: e",
    );
    expect(await openZip(bytes).read("e")).toEqual(new Uint8Array([1]));
  }, 30000);

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
      "Unable to decompress this ZIP entry with deflate-raw.",
    );
  });

  it("does not apply the deflate cap to stored entries", async () => {
    const bytes = new Uint8Array(zipSync({ "file.txt": strToU8("content") }, { level: 0 }));
    const central = findSignature(bytes, 0x02014b50);
    new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).setUint32(
      central + 24,
      64 * 1024 * 1024 + 1,
      true,
    );
    await expect(openZip(bytes).read("file.txt")).rejects.toThrow(
      "Invalid ZIP entry size: file.txt",
    );
  });

  it("rejects deflated entries over the 64 MiB limit", async () => {
    const bytes = new Uint8Array(zipSync({ "file.txt": strToU8("content") }, { level: 9 }));
    const central = findSignature(bytes, 0x02014b50);
    new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).setUint32(
      central + 24,
      64 * 1024 * 1024 + 1,
      true,
    );

    await expect(openZip(bytes).read("file.txt")).rejects.toThrow(
      "ZIP entry exceeds the decompressed size limit: file.txt",
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

const findSignature = (bytes: Uint8Array, signature: number): number => {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  for (let offset = 0; offset <= bytes.byteLength - 4; offset++)
    if (view.getUint32(offset, true) === signature) return offset;
  throw new Error("Signature not found.");
};
