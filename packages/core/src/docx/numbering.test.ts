import { describe, expect, it } from "vite-plus/test";
import { assignListMarkers, parseNumbering } from "./numbering";
import type { Block } from "./model";

describe("DOCX numbering", () => {
  it("formats levels and resets nested counters when the parent advances", () => {
    const doc = new DOMParser().parseFromString(
      `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:abstractNum w:abstractNumId="0">
          <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
          <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1.%2)"/></w:lvl>
        </w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
      </w:numbering>`,
      "application/xml",
    );
    const numbering = parseNumbering(doc);
    expect(numbering.marker("1", 0).text).toBe("1.");
    expect(numbering.marker("1", 0).text).toBe("1.");
    const counters = [3];
    expect(numbering.marker("1", 0, counters).counters).toEqual([4]);
    expect(counters).toEqual([3]);
    const blocks: Block[] = [0, 1, 1, 0, 1].map((level) => ({
      type: "paragraph",
      properties: {},
      runs: [],
      list: { numId: "1", level },
    }));
    expect(
      assignListMarkers(blocks, numbering).map((block) =>
        block.type === "paragraph" ? block.list?.marker : undefined,
      ),
    ).toEqual(["1.", "1.a)", "1.b)", "2.", "2.a)"]);
  });

  it("supports roman and letter number formats", () => {
    const doc = new DOMParser().parseFromString(
      `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="4"/><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="%1."/></w:lvl><w:lvl w:ilvl="1"><w:numFmt w:val="upperLetter"/><w:lvlText w:val="%2."/></w:lvl></w:abstractNum><w:num w:numId="2"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
      "application/xml",
    );
    const numbering = parseNumbering(doc);
    expect(numbering.marker("2", 0).text).toBe("iv.");
    expect(numbering.marker("2", 1).text).toBe("A.");
  });

  it.each([
    ["upperRoman", 9, "IX."],
    ["upperRoman", 1, "I."],
    ["upperRoman", 3999, "MMMCMXCIX."],
    ["upperRoman", 0, "0."],
    ["upperRoman", -1, "-1."],
    ["upperRoman", 4000, "4000."],
    ["upperRoman", 2000000000, "2000000000."],
    ["upperRoman", 1.5, "1.5."],
    ["upperRoman", Infinity, "Infinity."],
    ["upperRoman", -Infinity, "-Infinity."],
    ["upperRoman", NaN, "NaN."],
    ["upperLetter", 27, "AA."],
    ["lowerLetter", 26, "z."],
    ["decimalFullWidth", 12, "１２."],
    ["japaneseCounting", 12, "12."],
    ["unknown", 12, "12."],
  ])("formats %s markers", (format, start, expected) => {
    const doc = new DOMParser().parseFromString(
      `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="${start}"/><w:numFmt w:val="${format}"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
      "application/xml",
    );
    expect(parseNumbering(doc).marker("1", 0).text).toBe(expected);
  });

  it("uses concrete-level overrides and maps unknown bullet glyphs to the default", () => {
    const doc = new DOMParser().parseFromString(
      `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="4"/><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xf123;"/><w:pPr><w:ind w:start="400" w:hanging="200"/></w:pPr><w:rPr><w:rFonts w:hAnsi="Wingdings"/></w:rPr></w:lvl></w:lvlOverride></w:num></w:numbering>`,
      "application/xml",
    );
    const numbering = parseNumbering(doc);

    expect(numbering.level("1", 0)).toMatchObject({
      start: 4,
      left: 20,
      hanging: 10,
      font: "Wingdings",
    });
    expect(numbering.marker("1", 0).text).toBe("•");
    expect(numbering.marker("missing", 2).text).toBe("1.");
  });

  it("maps Word bullet glyphs and honors simple suffixes", () => {
    const doc = new DOMParser().parseFromString(
      `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xf0a7;"/><w:suff w:val="space"/></w:lvl><w:lvl w:ilvl="1"><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xf076;"/></w:lvl><w:lvl w:ilvl="2"><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xf0d8;"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
      "application/xml",
    );
    const numbering = parseNumbering(doc);
    expect(numbering.marker("1", 0)).toMatchObject({ text: "▪", suffix: "space" });
    expect(numbering.marker("1", 1).text).toBe("❖");
    expect(numbering.marker("1", 2).text).toBe("➢");
  });

  it("assigns independent counters in a fresh numbering fold", () => {
    const doc = new DOMParser().parseFromString(
      `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
      "application/xml",
    );
    const numbering = parseNumbering(doc);
    const blocks: Block[] = [
      { type: "paragraph", properties: {}, runs: [], list: { numId: "1", level: 0 } },
      { type: "paragraph", properties: {}, runs: [], list: { numId: "1", level: 0 } },
    ];
    expect(
      assignListMarkers(blocks, numbering).map((block) =>
        block.type === "paragraph" ? block.list?.marker : undefined,
      ),
    ).toEqual(["1.", "2."]);
    expect(
      assignListMarkers(blocks, numbering).map((block) =>
        block.type === "paragraph" ? block.list?.marker : undefined,
      ),
    ).toEqual(["1.", "2."]);
  });
});
