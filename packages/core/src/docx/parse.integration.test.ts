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

  it("parses comment references nested inside runs", async () => {
    const data = buildDocx({
      document:
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:commentReference w:id="0"/></w:r></w:p><w:sectPr/></w:body></w:document>',
    });
    const doc = await parseDocx(data);
    const paragraph = doc.sections[0].blocks[0];
    expect(paragraph.type).toBe("paragraph");
    if (paragraph.type !== "paragraph") return;
    expect(paragraph.runs).toEqual([{ type: "comment-reference", id: "0" }]);
  });

  it("inherits the last section properties for trailing blocks without body sectPr", async () => {
    const doc = await parseDocx(
      buildDocx({
        document:
          '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:sectPr><w:pgSz w:w="10000" w:h="12000"/><w:pgMar w:top="100" w:right="200" w:bottom="300" w:left="400"/></w:sectPr></w:pPr><w:r><w:t>first section</w:t></w:r></w:p><w:p><w:r><w:t>trailing block</w:t></w:r></w:p></w:body></w:document>',
      }),
    );
    expect(doc.sections).toHaveLength(2);
    expect(doc.sections[1].properties.width).toBe(doc.sections[0].properties.width);
    expect(doc.sections[1].properties.height).toBe(doc.sections[0].properties.height);
    expect(doc.sections[1].properties.margins).toEqual(doc.sections[0].properties.margins);
    expect(doc.sections[1].blocks[0]).toMatchObject({
      type: "paragraph",
      runs: [{ type: "text", text: "trailing block" }],
    });
  });

  it("resolves the main part and related Word parts from OPC relationships", async () => {
    const doc = await parseDocx(
      buildDocx({
        mainPath: "custom/doc.xml",
        stylesPath: "custom/styles.xml",
        commentsPath: "custom/review.xml",
        themePath: "custom/theme.xml",
        document:
          '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:commentRangeStart w:id="0"/><w:r><w:t>custom main</w:t></w:r><w:commentRangeEnd w:id="0"/></w:p><w:sectPr/></w:body></w:document>',
        styles:
          '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault></w:docDefaults></w:styles>',
        comments:
          '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="0" w:author="Reviewer"><w:p><w:r><w:t>custom comment</w:t></w:r></w:p></w:comment></w:comments>',
        theme:
          '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme><a:majorFont><a:latin typeface="Major"/></a:majorFont><a:minorFont><a:latin typeface="Minor"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>',
      }),
    );
    expect(doc.sections[0].blocks[0]).toMatchObject({
      type: "paragraph",
      runs: [
        { type: "text", text: "custom main", comments: ["0"], style: { fontFamily: "Minor" } },
      ],
    });
    expect(doc.comments[0]).toMatchObject({ author: "Reviewer" });
  });

  it("uses docGrid line pitch only for grid line modes", async () => {
    const make = (type: string) =>
      parseDocx(
        buildDocx({
          document: `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>grid</w:t></w:r></w:p><w:sectPr><w:docGrid w:type="${type}" w:linePitch="360"/></w:sectPr></w:body></w:document>`,
        }),
      );
    expect((await make("lines")).sections[0].properties.linePitch).toBe(18);
    expect((await make("linesAndChars")).sections[0].properties.linePitch).toBe(18);
    expect((await make("default")).sections[0].properties.linePitch).toBeUndefined();
  });
});
