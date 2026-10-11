import { describe, expect, it } from "vite-plus/test";
import { buildDocx } from "../../test/docx";
import { parseDocx } from "./parse";

const xml = (body: string) =>
  `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`;

describe("DOCX parser integration", () => {
  it("associates comment ranges with text and resolves replies", async () => {
    const doc = await parseDocx(
      buildDocx({
        document: xml(
          '<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Commented text</w:t></w:r><w:commentRangeEnd w:id="0"/></w:p>',
        ),
        comments:
          '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:comment w:id="0" w:author="Aiko"><w:p w14:paraId="AA"><w:r><w:t>Parent</w:t></w:r></w:p></w:comment><w:comment w:id="1" w:author="Morgan"><w:p w14:paraId="BB"><w:r><w:t>Reply</w:t></w:r></w:p></w:comment></w:comments>',
        commentsExtended:
          '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w15:commentEx w15:paraId="BB" w15:paraIdParent="AA"/></w15:commentsEx>',
      }),
    );
    const paragraph = doc.sections[0].blocks[0];
    expect(paragraph.type).toBe("paragraph");
    if (paragraph.type !== "paragraph") return;
    expect(paragraph.runs).toContainEqual(
      expect.objectContaining({ type: "text", text: "Commented text", comments: ["0"] }),
    );
    expect(doc.comments[1]).toMatchObject({ id: "1", parentId: "0" });
  });

  it("preserves inserted and deleted text with their authors", async () => {
    const doc = await parseDocx(
      buildDocx({
        document: xml(
          '<w:p><w:del w:author="Aiko"><w:r><w:delText>old</w:delText></w:r></w:del><w:ins w:author="Morgan"><w:r><w:t>new</w:t></w:r></w:ins></w:p>',
        ),
      }),
    );
    const paragraph = doc.sections[0].blocks[0];
    expect(paragraph.type).toBe("paragraph");
    if (paragraph.type !== "paragraph") return;
    const deleted = paragraph.runs.find((run) => run.type === "text" && run.text === "old");
    const inserted = paragraph.runs.find((run) => run.type === "text" && run.text === "new");
    expect(deleted?.type === "text" ? deleted.change : undefined).toMatchObject({
      kind: "del",
      author: "Aiko",
    });
    expect(inserted?.type === "text" ? inserted.change : undefined).toMatchObject({
      kind: "ins",
      author: "Morgan",
    });
  });

  it("resolves grid spans and vertical merges by grid column", async () => {
    const doc = await parseDocx(
      buildDocx({
        document: xml(
          '<w:tbl><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>span</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>top</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>span again</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p><w:r><w:t>bottom</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
        ),
      }),
    );
    const table = doc.sections[0].blocks[0];
    expect(table.type).toBe("table");
    if (table.type !== "table") return;
    expect(table.rows[0].cells[0].gridSpan).toBe(2);
    expect(table.rows[0].cells[1].rowSpan).toBe(2);
  });

  it("emits a PAGE field from a simple field instruction", async () => {
    const doc = await parseDocx(
      buildDocx({
        document: xml(
          '<w:p><w:fldSimple w:instr=" PAGE "><w:r><w:rPr><w:sz w:val="28"/></w:rPr><w:t>3</w:t></w:r></w:fldSimple></w:p>',
        ),
      }),
    );
    const paragraph = doc.sections[0].blocks[0];
    expect(paragraph.type).toBe("paragraph");
    if (paragraph.type !== "paragraph") return;
    const field = paragraph.runs.find((run) => run.type === "field");
    expect(field).toMatchObject({ type: "field", name: "PAGE" });
    expect(field?.type === "field" ? field.style?.fontSize : undefined).toBe(14);
  });
});
