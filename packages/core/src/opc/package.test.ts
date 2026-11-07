import { describe, expect, it } from "vite-plus/test";
import { strToU8, zipSync } from "fflate";
import { buildDocx } from "../../test/docx";
import { openOpc } from "./package";

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
});
