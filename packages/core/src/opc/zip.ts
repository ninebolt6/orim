const decoder = new TextDecoder();
const MAX_ENTRY_SIZE = 64 * 1024 * 1024;

const bomEncoding = (bytes: Uint8Array): string | undefined => {
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return "utf-16le";
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return "utf-16be";
  return undefined;
};

const xmlSniffedEncoding = (bytes: Uint8Array): string => {
  const bom = bomEncoding(bytes);
  if (bom) return bom;
  if (bytes[0] === 0x3c && bytes[1] === 0 && bytes[2] === 0x3f && bytes[3] === 0) return "utf-16le";
  if (bytes[0] === 0 && bytes[1] === 0x3c && bytes[2] === 0 && bytes[3] === 0x3f) return "utf-16be";
  return "utf-8";
};

const decodeText = (bytes: Uint8Array): string => {
  const sniffed = xmlSniffedEncoding(bytes);
  const text = new TextDecoder(sniffed).decode(bytes);
  const declared = /^<\?xml\s[^?]*?\bencoding\s*=\s*["']([^"']+)["']/i.exec(text)?.[1];
  if (!declared || (sniffed.startsWith("utf-16") && /^utf-?16(?:le|be)?$/i.test(declared)))
    return text;
  try {
    const declaredDecoder = new TextDecoder(declared);
    return declaredDecoder.encoding === sniffed ? text : declaredDecoder.decode(bytes);
  } catch {
    return text;
  }
};

const u16 = (view: DataView, offset: number): number => {
  return view.getUint16(offset, true);
};

const u32 = (view: DataView, offset: number): number => {
  return view.getUint32(offset, true);
};

const inflateRaw = async (bytes: Uint8Array, declaredSize: number): Promise<Uint8Array> => {
  if (typeof DecompressionStream === "undefined") {
    throw new Error("This browser does not support ZIP deflate decompression.");
  }
  const chunks: Uint8Array[] = [];
  let total = 0;
  const limit = Math.min(declaredSize, MAX_ENTRY_SIZE);
  let reader: ReadableStreamDefaultReader<Uint8Array> | undefined;
  try {
    const stream = new Blob([new Uint8Array(bytes).buffer])
      .stream()
      .pipeThrough(new DecompressionStream("deflate-raw" as CompressionFormat));
    reader = stream.getReader();
    for (;;) {
      const result = await reader.read();
      if (result.done) break;
      total += result.value.byteLength;
      if (total > limit) throw new Error();
      chunks.push(result.value);
    }
  } catch {
    reader?.cancel().catch(() => undefined);
    throw new Error("Unable to decompress this ZIP entry with deflate-raw.");
  }
  const output = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    output.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return output;
};

export interface ZipArchive {
  has(path: string): boolean;
  read(path: string): Promise<Uint8Array>;
  text(path: string): Promise<string>;
  paths(): string[];
}

export const openZip = (input: ArrayBuffer | Uint8Array): ZipArchive => {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  if (
    [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1].every((byte, index) => bytes[index] === byte)
  )
    throw new Error(
      "Unsupported file: this is a legacy binary Office file (.doc/.xls/.ppt) or a password-protected document. Save it as .docx/.xlsx/.pptx without a password.",
    );
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const minOffset = Math.max(0, bytes.byteLength - 22 - 0xffff);
  let end = -1;
  for (let i = bytes.byteLength - 22; i >= minOffset; i--) {
    if (u32(view, i) === 0x06054b50) {
      end = i;
      break;
    }
  }
  if (end < 0) throw new Error("Invalid ZIP: end-of-central-directory record not found.");
  const disk = u16(view, end + 4);
  const directoryDisk = u16(view, end + 6);
  const count = u16(view, end + 10);
  const directorySize = u32(view, end + 12);
  const directoryOffset = u32(view, end + 16);
  if (disk || directoryDisk || count === 0xffff || directoryOffset === 0xffffffff) {
    throw new Error("ZIP64 and multi-disk archives are not supported.");
  }
  if (directoryOffset + directorySize > bytes.byteLength) {
    throw new Error("Invalid ZIP: central directory is outside the archive.");
  }

  const entries = new Map<
    string,
    { method: number; size: number; compressedSize: number; offset: number }
  >();
  let cursor = directoryOffset;
  for (let i = 0; i < count; i++) {
    if (u32(view, cursor) !== 0x02014b50) throw new Error("Invalid ZIP central directory.");
    const flags = u16(view, cursor + 8);
    if (flags & 1) throw new Error("Encrypted ZIP entries are not supported.");
    const method = u16(view, cursor + 10);
    const compressedSize = u32(view, cursor + 20);
    const size = u32(view, cursor + 24);
    const nameLength = u16(view, cursor + 28);
    const extraLength = u16(view, cursor + 30);
    const commentLength = u16(view, cursor + 32);
    const localOffset = u32(view, cursor + 42);
    if (size === 0xffffffff || compressedSize === 0xffffffff || localOffset === 0xffffffff) {
      throw new Error("ZIP64 archives are not supported.");
    }
    const name = decoder.decode(bytes.subarray(cursor + 46, cursor + 46 + nameLength));
    entries.set(name, { method, size, compressedSize, offset: localOffset });
    cursor += 46 + nameLength + extraLength + commentLength;
  }

  let remaining = Math.max(256 * 1024 * 1024, bytes.byteLength * 8);
  return {
    has: (path) => entries.has(path),
    paths: () => [...entries.keys()],
    async read(path) {
      const entry = entries.get(path);
      if (!entry) throw new Error(`ZIP entry not found: ${path}`);
      if ((entry.method === 8 && entry.size > MAX_ENTRY_SIZE) || entry.size > remaining)
        throw new Error(`ZIP entry exceeds the decompressed size limit: ${path}`);
      remaining -= entry.size;
      const local = entry.offset;
      if (u32(view, local) !== 0x04034b50) throw new Error(`Invalid ZIP local header: ${path}`);
      const nameLength = u16(view, local + 26);
      const extraLength = u16(view, local + 28);
      const start = local + 30 + nameLength + extraLength;
      const compressed = bytes.subarray(start, start + entry.compressedSize);
      let content: Uint8Array;
      if (entry.method === 0) content = compressed.slice();
      else if (entry.method === 8) content = await inflateRaw(compressed, entry.size);
      else throw new Error(`Unsupported ZIP compression method ${entry.method}: ${path}`);
      if (content.byteLength !== entry.size) throw new Error(`Invalid ZIP entry size: ${path}`);
      return content;
    },
    async text(path) {
      return decodeText(await this.read(path));
    },
  };
};
