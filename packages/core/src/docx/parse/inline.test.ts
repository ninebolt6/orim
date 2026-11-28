import { describe, expect, it } from "vite-plus/test";
import { parseContext, wordElement } from "../../../test/docx-context";
import { collectInline } from "./inline";
import { createCommentRangeState, MAX_OPEN_COMMENT_RANGES } from "./context";

describe("inline parser", () => {
  it("parses text atoms, comment ranges, bookmarks, and hyperlink targets", () => {
    const paragraph = wordElement(
      '<w:p><w:commentRangeStart w:id="0"/><w:r><w:tab/><w:sym w:char="0041"/><w:sym w:char="invalid"/><w:noBreakHyphen/><w:softHyphen/><w:t>text</w:t><w:delInstrText>deleted instruction</w:delInstrText></w:r><w:commentRangeEnd w:id="0"/><w:bookmarkStart w:name="mark"/><w:hyperlink w:anchor="section"><w:r><w:t>anchor</w:t></w:r></w:hyperlink><w:hyperlink><w:r><w:t>plain</w:t></w:r></w:hyperlink><w:proofErr><w:r><w:t>ignored</w:t></w:r></w:proofErr><w:smartTag><w:r><w:t>nested</w:t></w:r></w:smartTag></w:p>',
    );
    const result = collectInline(paragraph, parseContext());

    expect(result).toMatchObject([
      { type: "text", text: "\t" },
      { type: "text", text: "A" },
      { type: "text", text: "‑" },
      { type: "text", text: "\u00ad" },
      { type: "text", text: "text", comments: ["0"] },
      { type: "text", text: "deleted instruction", comments: ["0"] },
      { type: "bookmark", name: "mark" },
      { type: "text", text: "anchor", href: "#section" },
      { type: "text", text: "plain" },
      { type: "text", text: "nested" },
    ]);
  });

  it("resolves image, relationship, field-character, and simple hyperlink instructions", () => {
    const context = parseContext({
      rels: new Map([
        [
          "image",
          {
            id: "image",
            type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
            target: "media/image.png",
            external: false,
          },
        ],
        [
          "link",
          {
            id: "link",
            type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            target: "https://example.com",
            external: true,
          },
        ],
      ]),
    });
    const paragraph = wordElement(
      '<w:p><w:r><w:drawing><wp:extent cx="12700" cy="25400"/><a:blip r:embed="image"/></w:drawing></w:r><w:hyperlink r:id="link"><w:r><w:t>related</w:t></w:r></w:hyperlink><w:fldSimple w:instr=" PAGE "><w:r><w:rPr><w:sz w:val="28"/></w:rPr><w:t>1</w:t></w:r></w:fldSimple><w:fldSimple w:instr=" HYPERLINK &quot;https://example.org&quot; "><w:r><w:t>field link</w:t></w:r></w:fldSimple></w:p>',
    );
    const result = collectInline(paragraph, context);

    expect(result).toContainEqual(
      expect.objectContaining({
        type: "image",
        path: "word/media/image.png",
        width: 1,
        height: 2,
      }),
    );
    expect(result).toContainEqual(
      expect.objectContaining({ type: "text", text: "related", href: "https://example.com" }),
    );
    expect(result).toContainEqual(
      expect.objectContaining({ type: "field", name: "PAGE", style: { fontSize: 14 } }),
    );
    expect(result).toContainEqual(
      expect.objectContaining({
        type: "text",
        text: "field link",
        href: "https://example.org",
      }),
    );
  });

  it("falls back to AlternateContent and emits only live page fields", () => {
    const paragraph = wordElement(
      '<w:p><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice><w:r><w:t>choice</w:t></w:r></mc:Choice><mc:Fallback><w:r><w:t>fallback</w:t></w:r></mc:Fallback></mc:AlternateContent><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> DATE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>cached</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:br w:type="page"/><w:cr/></w:p>',
    );
    const result = collectInline(paragraph, parseContext());

    expect(result).toContainEqual(
      expect.objectContaining({ type: "text", text: "fallback", comments: [] }),
    );
    expect(result).toContainEqual({ type: "break", page: true });
    expect(result).toContainEqual({ type: "break", page: false });
    expect(result.some((inline) => inline.type === "text" && inline.text === "choice")).toBe(false);
    expect(result.some((inline) => inline.type === "text" && inline.text === "cached")).toBe(true);
    expect(result.some((inline) => inline.type === "field")).toBe(false);
  });

  it("bounds simultaneously open comment ranges", () => {
    const ranges = createCommentRangeState();
    for (let index = 0; index < MAX_OPEN_COMMENT_RANGES + 8; index++) ranges.start(String(index));

    expect(ranges.snapshot()).toHaveLength(MAX_OPEN_COMMENT_RANGES);
    ranges.end("0");
    ranges.start("later");
    expect(ranges.snapshot()).toContain("later");
  });
});
