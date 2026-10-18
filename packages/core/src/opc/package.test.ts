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

  it("rejects malformed XML and bounds internal entity expansion", async () => {
    const malformed = openOpc(zipSync({ "malformed.xml": strToU8("<root>") }));
    await expect(malformed.xml("malformed.xml")).rejects.toThrow("Invalid XML part");

    const declarations = ['<!ENTITY e0 "x">'];
    for (let index = 1; index < 10; index++)
      declarations.push(
        `<!ENTITY e${index} "${Array.from({ length: 10 }, () => `&e${index - 1};`).join("")}">`,
      );
    const xml = `<!DOCTYPE root [${declarations.join("")}]><root>&e7;</root>`;
    const expanded = openOpc(zipSync({ "expanded.xml": strToU8(xml) }));
    try {
      const document = await expanded.xml("expanded.xml");
      expect(document.documentElement.textContent?.length ?? 0).toBeLessThan(1_000_000);
    } catch (error) {
      expect(error).toBeInstanceOf(Error);
    }
  });
});
