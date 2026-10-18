import { describe, expect, it } from "vite-plus/test";
import { buildDocx } from "../../test/docx";
import { parseDocx } from "./parse";

describe("DOCX package parsing", () => {
  it("inherits the last section properties for trailing blocks without body sectPr", async () => {
    const doc = await parseDocx(
      buildDocx({
        header:
          '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>inherited header</w:t></w:r></w:p></w:hdr>',
        document:
          '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rHeader"/><w:pgSz w:w="10000" w:h="12000"/><w:pgMar w:top="100" w:right="200" w:bottom="300" w:left="400"/></w:sectPr></w:pPr><w:r><w:t>first section</w:t></w:r></w:p><w:p><w:r><w:t>trailing block</w:t></w:r></w:p></w:body></w:document>',
      }),
    );
    expect(doc.sections).toHaveLength(2);
    expect(doc.sections[1].properties.width).toBe(doc.sections[0].properties.width);
    expect(doc.sections[1].properties.height).toBe(doc.sections[0].properties.height);
    expect(doc.sections[1].properties.margins).toEqual(doc.sections[0].properties.margins);
    expect(doc.sections[1].properties.header?.[0]).toMatchObject({
      type: "paragraph",
      runs: [{ type: "text", text: "inherited header" }],
    });
    expect(doc.sections[1].blocks[0]).toMatchObject({
      type: "paragraph",
      runs: [{ type: "text", text: "trailing block" }],
    });
  });

  it("keeps body numbering continuous across sections and resets header stories", async () => {
    const listParagraph = (text: string) =>
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
    const list = `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
    const doc = await parseDocx(
      buildDocx({
        numbering: list,
        header: `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${listParagraph("header one")}${listParagraph("header two")}</w:hdr>`,
        document: `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:sectPr><w:headerReference w:type="default" r:id="rHeader"/><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr><w:r><w:t>body one</w:t></w:r></w:p>${listParagraph("body two")}<w:sectPr/></w:body></w:document>`,
      }),
    );
    const markers = (blocks: (typeof doc.sections)[number]["blocks"]) =>
      blocks.flatMap((block) =>
        block.type === "paragraph" && block.list?.marker ? [block.list.marker] : [],
      );
    expect(doc.sections).toHaveLength(2);
    expect(markers(doc.sections.flatMap((section) => section.blocks))).toEqual(["1.", "2."]);
    expect(markers(doc.sections[0].properties.header ?? [])).toEqual(["1.", "2."]);
    expect(markers(doc.sections[1].properties.header ?? [])).toEqual(["1.", "2."]);
  });

  it("resolves the main part and related Word parts from OPC relationships", async () => {
    const doc = await parseDocx(
      buildDocx({
        mainPath: "custom/doc.xml",
        stylesPath: "custom/styles.xml",
        numberingPath: "custom/numbering.xml",
        commentsPath: "custom/review.xml",
        commentsExtendedPath: "custom/reviewExtended.xml",
        themePath: "custom/theme.xml",
        document:
          '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:commentRangeStart w:id="0"/><w:r><w:t>custom main</w:t></w:r><w:commentRangeEnd w:id="0"/></w:p><w:sectPr/></w:body></w:document>',
        styles:
          '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault></w:docDefaults></w:styles>',
        numbering:
          '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>',
        comments:
          '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:comment w:id="0" w:author="Reviewer"><w:p w14:paraId="AA"><w:r><w:t>custom comment</w:t></w:r></w:p></w:comment><w:comment w:id="1" w:author="Reply"><w:p w14:paraId="BB"><w:r><w:t>custom reply</w:t></w:r></w:p></w:comment></w:comments>',
        commentsExtended:
          '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w15:commentEx w15:paraId="BB" w15:paraIdParent="AA"/></w15:commentsEx>',
        theme:
          '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme><a:majorFont><a:latin typeface="Major"/></a:majorFont><a:minorFont><a:latin typeface="Minor"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>',
      }),
    );
    expect(doc.sections[0].blocks[0]).toMatchObject({
      type: "paragraph",
      runs: [
        { type: "text", text: "custom main", comments: ["0"], style: { fontFamily: "Minor" } },
      ],
      list: { marker: "1." },
    });
    expect(doc.comments[0]).toMatchObject({ author: "Reviewer" });
    expect(doc.comments[1]).toMatchObject({ parentId: "0" });
  });
});
