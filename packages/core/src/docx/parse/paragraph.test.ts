import { describe, expect, it } from "vite-plus/test";
import { parseContext, wordElement } from "../../../test/docx-context";
import { parseParagraph } from "./paragraph";

describe("paragraph parser", () => {
  it("associates comment ranges with text", () => {
    const paragraph = parseParagraph(
      wordElement(
        '<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Commented text</w:t></w:r><w:commentRangeEnd w:id="0"/></w:p>',
      ),
      parseContext(),
    );

    expect(paragraph.runs).toContainEqual(
      expect.objectContaining({
        type: "text",
        text: "Commented text",
        comments: ["0"],
      }),
    );
  });

  it("preserves inserted and deleted text with their authors", () => {
    const paragraph = parseParagraph(
      wordElement(
        '<w:p><w:del w:author="Aiko"><w:r><w:delText>old</w:delText></w:r></w:del><w:ins w:author="Morgan"><w:r><w:t>new</w:t></w:r></w:ins></w:p>',
      ),
      parseContext(),
    );

    expect(paragraph.runs).toContainEqual(
      expect.objectContaining({
        type: "text",
        text: "old",
        change: { kind: "del", author: "Aiko", date: "" },
      }),
    );
    expect(paragraph.runs).toContainEqual(
      expect.objectContaining({
        type: "text",
        text: "new",
        change: { kind: "ins", author: "Morgan", date: "" },
      }),
    );
  });

  it("keeps a run-level deletion separate from paragraph-mark revisions", () => {
    const paragraph = parseParagraph(
      wordElement(
        '<w:p><w:del w:author="Aiko"><w:r><w:delText>deleted run</w:delText></w:r></w:del></w:p>',
      ),
      parseContext(),
    );

    expect(paragraph.properties.change).toBeUndefined();
    expect(paragraph.runs).toContainEqual(
      expect.objectContaining({
        type: "text",
        text: "deleted run",
        change: { kind: "del", author: "Aiko", date: "" },
      }),
    );
  });

  it("keeps inserted run properties scoped to the run", () => {
    const paragraph = parseParagraph(
      wordElement(
        '<w:p><w:ins w:author="Morgan"><w:r><w:rPr><w:b/></w:rPr><w:t>inserted run</w:t></w:r></w:ins></w:p>',
      ),
      parseContext(),
    );

    expect(paragraph.properties.change).toBeUndefined();
    expect(paragraph.runs).toContainEqual(
      expect.objectContaining({
        type: "text",
        text: "inserted run",
        change: { kind: "ins", author: "Morgan", date: "" },
        style: expect.objectContaining({ bold: true }),
      }),
    );
  });

  it("parses comment references nested inside runs", () => {
    const paragraph = parseParagraph(
      wordElement('<w:p><w:r><w:commentReference w:id="0"/></w:r></w:p>'),
      parseContext(),
    );

    expect(paragraph.runs).toEqual([{ type: "comment-reference", id: "0" }]);
  });

  it("falls back to level zero for an out-of-range paragraph list level", () => {
    const paragraph = parseParagraph(
      wordElement(
        '<w:p><w:pPr><w:numPr><w:ilvl w:val="999999999999"/><w:numId w:val="1"/></w:numPr></w:pPr></w:p>',
      ),
      parseContext(),
    );

    expect(paragraph.list?.level).toBe(0);
  });

  it("emits a PAGE field from a simple field instruction", () => {
    const paragraph = parseParagraph(
      wordElement(
        '<w:p><w:fldSimple w:instr=" PAGE "><w:r><w:rPr><w:sz w:val="28"/></w:rPr><w:t>3</w:t></w:r></w:fldSimple></w:p>',
      ),
      parseContext(),
    );

    expect(paragraph.runs).toContainEqual(
      expect.objectContaining({ type: "field", name: "PAGE", style: { fontSize: 14 } }),
    );
  });
});
