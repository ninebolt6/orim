import { describe, expect, it } from "vite-plus/test";
import { strToU8, zipSync } from "fflate";
import { buildDocx } from "../../test/docx";
import { openOpc, readMedia } from "./package";

describe("OPC relationships", () => {
  it("resolves relative and package-absolute parts and preserves external relationships", async () => {
    const opc = openOpc(
      buildDocx({
        document:
          '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>',
        relationships:
          '<Relationship Id="external" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>',
      }),
    );
    expect(opc.resolve("word/document.xml", "styles.xml")).toBe("word/styles.xml");
    expect(opc.resolve("word/document.xml", "/word/media/image.png")).toBe("word/media/image.png");
    expect((await opc.relationships("word/document.xml")).get("external")).toEqual(
      expect.objectContaining({ target: "https://example.com", external: true }),
    );
  });

  it("rejects malformed XML", async () => {
    const malformed = openOpc(zipSync({ "malformed.xml": strToU8("<root>") }));
    await expect(malformed.xml("malformed.xml")).rejects.toThrow("Invalid XML part");
  });

  it.each(["<!DOCTYPE root>", '<!DOCTYPE root [<!ENTITY e "text">]>'])(
    "rejects XML parts with DTD %s",
    async (doctype) => {
      const opc = openOpc(zipSync({ "part.xml": strToU8(`${doctype}<root/>`) }));
      await expect(opc.xml("part.xml")).rejects.toThrow("Invalid XML part: part.xml");
    },
  );

  it.each(["<!DOCTYPE Types><Types/>", "<Types>"])(
    "rejects invalid content types XML %s",
    async (xml) => {
      const opc = openOpc(zipSync({ "[Content_Types].xml": strToU8(xml) }));
      await expect(opc.mime("image.png")).rejects.toThrow("Invalid XML part: [Content_Types].xml");
    },
  );

  it.each([
    ["single-quoted", "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"],
    ["double-quoted", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'],
    ["BOM-prefixed", '\uFEFF<?xml version="1.0"?>'],
  ])("parses parts with a %s XML declaration", async (_label, declaration) => {
    const opc = openOpc(zipSync({ "part.xml": strToU8(`${declaration}<root ok="1"/>`) }));
    expect((await opc.xml("part.xml")).documentElement.getAttribute("ok")).toBe("1");
  });

  it.each([
    ["UTF-16LE with a BOM", utf16Bytes('<root ok="1">😀</root>', "le", true)],
    ["UTF-16BE with a BOM", utf16Bytes('<root ok="1">😀</root>', "be", true)],
    [
      "UTF-16LE with an encoding declaration",
      utf16Bytes('<?xml version="1.0" encoding="UTF-16"?><root ok="1">😀</root>', "le", false),
    ],
  ])("parses a %s XML part", async (_label, bytes) => {
    const opc = openOpc(zipSync({ "part.xml": bytes }));
    const root = (await opc.xml("part.xml")).documentElement;
    expect(root.getAttribute("ok")).toBe("1");
    expect(root.textContent).toBe("😀");
  });

  it.each([
    ["csUTF16", "le"],
    ["csUTF16LE", "le"],
    ["csUTF16BE", "be"],
  ] as const)("parses UTF-16 parts declaring the %s alias", async (encoding, endian) => {
    const bytes = utf16Bytes(
      `<?xml version="1.0" encoding="${encoding}"?><root ok="1"/>`,
      endian,
      true,
    );
    const opc = openOpc(zipSync({ "part.xml": bytes }));
    expect((await opc.xml("part.xml")).documentElement.getAttribute("ok")).toBe("1");
  });

  it.each([
    ["le", "UTF-16BE"],
    ["be", "UTF-16LE"],
  ] as const)("rejects a %s part declaring %s", async (endian, encoding) => {
    const bytes = utf16Bytes(`<?xml version="1.0" encoding="${encoding}"?><root/>`, endian, true);
    const opc = openOpc(zipSync({ "part.xml": bytes }));
    await expect(opc.xml("part.xml")).rejects.toThrow(`Unsupported XML encoding: ${encoding}`);
  });

  it("rejects UCS-2 declarations outside the UTF-16 encoding form", async () => {
    const bytes = utf16Bytes('<?xml version="1.0" encoding="ISO-10646-UCS-2"?><root/>', "le", true);
    const opc = openOpc(zipSync({ "part.xml": bytes }));
    await expect(opc.xml("part.xml")).rejects.toThrow("Unsupported XML encoding: ISO-10646-UCS-2");
  });

  it("rejects whitespace in encoding names", async () => {
    const bytes = utf16Bytes('<?xml version="1.0" encoding=" UCS-2 "?><root/>', "le", true);
    const opc = openOpc(zipSync({ "part.xml": bytes }));
    await expect(opc.xml("part.xml")).rejects.toThrow("Unsupported XML encoding:  UCS-2 ");
  });

  it.each(["utf8", "unicode-1-1-utf-8", "x-unicode20utf8", "csUTF8"])(
    "keeps UTF-8 parts readable when %s is declared",
    async (encoding) => {
      const opc = openOpc(
        zipSync({
          "part.xml": strToU8(`<?xml version="1.0" encoding="${encoding}"?><root ok="é"/>`),
        }),
      );
      expect((await opc.xml("part.xml")).documentElement.getAttribute("ok")).toBe("é");
    },
  );

  it("rejects non-Unicode encoding declarations", async () => {
    const opc = openOpc(
      zipSync({
        "part.xml": strToU8('<?xml version="1.0" encoding="windows-1252"?><root/>'),
      }),
    );
    await expect(opc.xml("part.xml")).rejects.toThrow("Unsupported XML encoding: windows-1252");
  });

  it.each([
    ["csUTF-16", "le"],
    ["csUTF-16LE", "le"],
    ["csUTF-16BE", "be"],
  ] as const)("rejects the nonstandard UTF-16 alias %s", async (encoding, endian) => {
    const bytes = utf16Bytes(`<?xml version="1.0" encoding="${encoding}"?><root/>`, endian, true);
    const opc = openOpc(zipSync({ "part.xml": bytes }));
    await expect(opc.xml("part.xml")).rejects.toThrow(`Unsupported XML encoding: ${encoding}`);
  });

  it("loads content types lazily so other part failures do not leak rejections", async () => {
    const opc = openOpc(
      zipSync({ "[Content_Types].xml": strToU8("<Types>"), "part.xml": strToU8("<root/>") }),
    );
    expect((await opc.xml("part.xml")).documentElement.localName).toBe("root");
  });

  it("skips media entries that are missing or unreadable", async () => {
    const opc = openOpc(zipSync({ "media/ok.png": new Uint8Array([1, 2, 3]) }));
    const media = await readMedia(opc, ["media/ok.png", "media/missing.png", "media/ok.png"]);
    expect(Object.keys(media)).toEqual(["media/ok.png"]);
    expect(media["media/ok.png"]!.mime).toBe("application/octet-stream");
  });
});

const utf16Bytes = (text: string, endian: "le" | "be", bom: boolean): Uint8Array => {
  const bytes = new Uint8Array(text.length * 2 + (bom ? 2 : 0));
  let offset = 0;
  if (bom) {
    bytes[offset++] = endian === "le" ? 0xff : 0xfe;
    bytes[offset++] = endian === "le" ? 0xfe : 0xff;
  }
  for (let index = 0; index < text.length; index++) {
    const code = text.charCodeAt(index);
    if (endian === "le") {
      bytes[offset++] = code & 0xff;
      bytes[offset++] = code >> 8;
    } else {
      bytes[offset++] = code >> 8;
      bytes[offset++] = code & 0xff;
    }
  }
  return bytes;
};
