import { afterEach, describe, expect, it, vi } from "vite-plus/test";
import { buildDocx } from "../../test/docx";
import { renderDocx } from "./document";

const documentXml = (content: string) =>
  `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>${content}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`;

describe("DOCX rendering integration", () => {
  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("reuses image object URLs across redraws and revokes them on destroy", async () => {
    vi.spyOn(URL, "createObjectURL").mockReturnValue("blob:orim-test");
    const revoke = vi.spyOn(URL, "revokeObjectURL").mockImplementation(() => {});
    const data = buildDocx({
      document: documentXml(
        '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="12700" cy="12700"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="image1"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>',
      ),
      relationships:
        '<Relationship Id="image1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image.png"/>',
      media: { "word/media/image.png": { bytes: new Uint8Array([1, 2, 3]), mime: "image/png" } },
    });
    const container = document.createElement("div");
    const rendered = await renderDocx(data, container);
    expect(container.querySelector("img")?.getAttribute("src")).toBe("blob:orim-test");
    rendered.update({ changes: "final" });
    rendered.update({ comments: false });
    expect(URL.createObjectURL).toHaveBeenCalledOnce();
    rendered.destroy();
    expect(revoke).toHaveBeenCalledOnce();
    expect(revoke).toHaveBeenCalledWith("blob:orim-test");
    expect(container.childElementCount).toBe(0);
  });

  it("removes and restores comment balloons and anchored highlights", async () => {
    const data = buildDocx({
      document: documentXml(
        '<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Commented text</w:t></w:r><w:commentRangeEnd w:id="0"/></w:p>',
      ),
      comments:
        '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="0" w:author="Aiko"><w:p><w:r><w:t>Review this text.</w:t></w:r></w:p></w:comment></w:comments>',
    });
    const container = document.createElement("div");
    const rendered = await renderDocx(data, container);
    expect(container.querySelector(".orim-comment")).not.toBeNull();
    expect(container.querySelector(".orim-comment-anchor")).not.toBeNull();
    rendered.update({ comments: false });
    expect(container.querySelector(".orim-comment")).toBeNull();
    expect(container.querySelector(".orim-comment-anchor")).toBeNull();
    rendered.update({ comments: true });
    expect(container.querySelector(".orim-comment")).not.toBeNull();
    expect(container.querySelector(".orim-comment-anchor")).not.toBeNull();
    rendered.destroy();
  });

  it("filters images in tracked changes by mode", async () => {
    vi.spyOn(URL, "createObjectURL").mockReturnValue("blob:revision-image");
    const drawing = (id: string) =>
      `<w:r><w:drawing><wp:inline><wp:extent cx="12700" cy="12700"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="${id}"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
    const data = buildDocx({
      document: documentXml(
        `<w:p><w:del w:author="A">${drawing("deleted")}</w:del><w:ins w:author="B">${drawing("inserted")}</w:ins></w:p>`,
      ),
      relationships:
        '<Relationship Id="deleted" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/deleted.png"/><Relationship Id="inserted" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/inserted.png"/>',
      media: {
        "word/media/deleted.png": { bytes: new Uint8Array([1]), mime: "image/png" },
        "word/media/inserted.png": { bytes: new Uint8Array([2]), mime: "image/png" },
      },
    });
    const container = document.createElement("div");
    const rendered = await renderDocx(data, container);
    expect(container.querySelectorAll("img")).toHaveLength(2);
    rendered.update({ changes: "final" });
    expect(container.querySelectorAll("img")).toHaveLength(1);
    expect(container.querySelector(".orim-change--ins")).not.toBeNull();
    rendered.update({ changes: "original" });
    expect(container.querySelectorAll("img")).toHaveLength(1);
    expect(container.querySelector(".orim-change--del")).not.toBeNull();
    rendered.destroy();
  });
});
