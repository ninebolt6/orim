import { describe, expect, it } from "vite-plus/test";
import { buildDocx } from "../../test/docx";
import { parseDocx } from "./parse";

describe("minimal DOCX integration", () => {
  it("unwraps block-level content controls and parses comment replies", async () => {
    const data = buildDocx({
      document:
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:sdt><w:sdtContent><w:p><w:r><w:t>TOC entry</w:t></w:r></w:p></w:sdtContent></w:sdt><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>',
      comments:
        '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="0" w:author="A"><w:p w14:paraId="AA" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:r><w:t>Parent</w:t></w:r></w:p></w:comment><w:comment w:id="1" w:author="B"><w:p w14:paraId="BB" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:r><w:t>Reply</w:t></w:r></w:p></w:comment></w:comments>',
      commentsExtended:
        '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w15:commentEx w15:paraId="BB" w15:paraIdParent="AA"/></w15:commentsEx>',
    });
    const doc = await parseDocx(data);
    const paragraph = doc.sections[0].blocks[0];
    expect(paragraph.type).toBe("paragraph");
    if (paragraph.type === "paragraph")
      expect(paragraph.runs[0]).toMatchObject({ text: "TOC entry" });
    expect(doc.comments[1]).toMatchObject({ id: "1", parentId: "0" });
  });
});
