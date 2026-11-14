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
