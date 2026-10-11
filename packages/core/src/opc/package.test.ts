import { describe, expect, it } from "vite-plus/test";
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
});
