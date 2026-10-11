import { mkdirSync, writeFileSync } from "node:fs";
import { deflateSync } from "node:zlib";
import { strToU8, zipSync } from "fflate";

function crc32(bytes) {
  let checksum = 0xffffffff;
  for (const byte of bytes) {
    checksum ^= byte;
    for (let bit = 0; bit < 8; bit++) checksum = (checksum >>> 1) ^ (0xedb88320 & -(checksum & 1));
  }
  return (checksum ^ 0xffffffff) >>> 0;
}

const esc = (text) =>
  String(text).replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
const xml = (body) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${body}`;
const textRun = (text, props = "") =>
  `<w:r>${props ? `<w:rPr>${props}</w:rPr>` : ""}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
const p = (text, props = "", style = "") =>
  `<w:p>${style ? `<w:pPr><w:pStyle w:val="${style}"/>${props}</w:pPr>` : props ? `<w:pPr>${props}</w:pPr>` : ""}${text}</w:p>`;
const paragraph = (text, style, props = "") => p(textRun(text), props, style);
const f = (type) => `<w:fldSimple w:instr="${type}"><w:r><w:t>1</w:t></w:r></w:fldSimple>`;
const cell = (text, props = "") =>
  `<w:tc><w:tcPr>${props}</w:tcPr>${paragraph(text, "Normal")}</w:tc>`;

function png() {
  const chunk = (type, content) => {
    const name = Buffer.from(type);
    const payload = Buffer.concat([name, Buffer.from(content)]);
    const out = Buffer.alloc(content.length + 12);
    out.writeUInt32BE(content.length, 0);
    payload.copy(out, 4);
    out.writeUInt32BE(crc32(payload), content.length + 8);
    return out;
  };
  const header = Buffer.alloc(13);
  header.writeUInt32BE(2, 0);
  header.writeUInt32BE(2, 4);
  header[8] = 8;
  header[9] = 6;
  const pixels = Uint8Array.from([
    0, 66, 113, 245, 255, 245, 167, 66, 255, 0, 24, 128, 56, 255, 255, 255, 255, 255,
  ]);
  return new Uint8Array(
    Buffer.concat([
      Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]),
      chunk("IHDR", header),
      chunk("IDAT", deflateSync(pixels)),
      chunk("IEND", new Uint8Array()),
    ]),
  );
}

const body = [];
body.push(paragraph("orim document rendering sample", "Title"));
body.push(paragraph("A compact browser-only DOCX rendering demo", "Subtitle"));
body.push(
  paragraph(
    "Document layout follows section page size, orientation, margins, and the typography carried by WordprocessingML styles.",
    "Normal",
  ),
);
body.push(
  paragraph(
    "日本語の文章も含むサンプル文書です。ページごとの余白、見出し、コメント、および変更履歴を確認できます。",
    "Normal",
  ),
);
body.push(
  p(
    textRun("A direct zero after-spacing value"),
    '<w:pStyle w:val="Normal"/><w:spacing w:after="0"/>',
  ),
);
body.push(
  p(textRun("A contextual-spacing paragraph"), '<w:pStyle w:val="Normal"/><w:contextualSpacing/>'),
);
body.push(
  p(
    textRun("A second contextual-spacing paragraph"),
    '<w:pStyle w:val="Normal"/><w:contextualSpacing/>',
  ),
);
body.push(paragraph("Overview", "Heading1"));
for (let i = 1; i <= 8; i++) {
  body.push(
    paragraph(
      `This introductory paragraph ${i} demonstrates body text, spacing, and a style inherited from Normal. Each paragraph is intentionally long enough to wrap naturally at the page content width. Word layout uses point-sized margins and font metrics in the browser.`,
      "BodyText",
    ),
  );
}
body.push(
  p(
    `<w:commentRangeStart w:id="0"/>${textRun("This note starts in one paragraph and continues into the next.")}`,
    "",
  ),
);
body.push(
  p(
    `${textRun("The second highlighted paragraph closes the comment range here.")}<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>`,
  ),
);
body.push(
  p(
    `${textRun("A tracked insertion: ")}<w:ins w:id="1" w:author="Aiko" w:date="2025-01-03T11:00:00Z">${textRun("new wording")}</w:ins>${textRun(" and a tracked deletion: ")}<w:del w:id="2" w:author="Morgan" w:date="2025-01-04T12:00:00Z"><w:r><w:delText>old wording</w:delText></w:r></w:del>`,
  ),
);
body.push(
  p(
    `${textRun("Moved text: ")}<w:moveFrom w:id="3" w:author="Morgan" w:date="2025-01-05T09:00:00Z">${textRun("before move")}</w:moveFrom><w:moveTo w:id="4" w:author="Morgan" w:date="2025-01-05T09:00:00Z">${textRun("after move")}</w:moveTo>`,
  ),
);
body.push(
  p(
    `<w:r><w:rPr><w:rPrChange w:id="5" w:author="Aiko" w:date="2025-01-06T10:00:00Z"><w:rPr><w:b/></w:rPr></w:rPrChange><w:i/></w:rPr><w:t>Formatting revision</w:t></w:r>`,
  ),
);
body.push(
  p(
    `${textRun("A paragraph with a deleted paragraph mark.")}`,
    '<w:pStyle w:val="Normal"/><w:del w:id="6" w:author="Morgan" w:date="2025-01-06T12:00:00Z"/>',
  ),
);
body.push(paragraph("Lists and tables", "Heading1"));
body.push(
  p(textRun("Numbered item one"), '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'),
);
body.push(
  p(
    textRun("Nested lower-letter item"),
    '<w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr>',
  ),
);
body.push(
  p(textRun("Numbered item two"), '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'),
);
body.push(p(textRun("Bullet item"), '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>'));
body.push(paragraph("A small embedded image and hyperlink:", "Normal"));
body.push(
  p(
    `<w:r><w:drawing><wp:inline><wp:extent cx="254000" cy="254000"/><wp:docPr id="1" name="sample"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rIdImage"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r><w:hyperlink r:id="rIdLink">${textRun("orim project")}</w:hyperlink>`,
  ),
);
body.push(
  paragraph(
    "The following table repeats its header on continuation pages and includes a vertically merged cell.",
    "Normal",
  ),
);
body.push(
  p(
    `${textRun("Before inline page break")}<w:r><w:br w:type="page"/></w:r>${textRun("After inline page break")}`,
  ),
);
body.push(
  `<w:sdt><w:sdtPr><w:docPartObj/></w:sdtPr><w:sdtContent>${p(
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${textRun("Contents")}`,
  )}${p(
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> HYPERLINK \\l "_TocOverview" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${textRun("Overview")}<w:r><w:fldChar w:fldCharType="end"/></w:r>${textRun(" 1")}`,
  )}${p(
    `${textRun("Second TOC entry")}<w:r><w:fldChar w:fldCharType="end"/></w:r>`,
  )}</w:sdtContent></w:sdt>`,
);
body.push(
  p(
    `<w:fldSimple w:instr="HYPERLINK &quot;https://example.org/orim&quot;">${textRun("Simple field hyperlink")}</w:fldSimple>`,
  ),
);
body.push(
  p(
    `${textRun("A paragraph with a bottom border.")}`,
    '<w:pStyle w:val="Normal"/><w:pBdr><w:bottom w:val="single" w:sz="6" w:color="auto" w:space="4"/></w:pBdr>',
  ),
);
body.push(
  p(`${textRun("Superscript example: ")}${textRun("2", '<w:vertAlign w:val="superscript"/>')}`),
);
body.push(
  p(
    `${textRun("Styled PAGE field: ")}<w:r><w:rPr><w:sz w:val="30"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:sz w:val="30"/></w:rPr><w:t>2</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`,
  ),
);
body.push(
  p(textRun("A numbered heading inherited through style."), '<w:pStyle w:val="NumberedHeading"/>'),
);
body.push(
  p(
    textRun("Direct list level overrides the style level."),
    '<w:pStyle w:val="NumberedHeading"/><w:numPr><w:ilvl w:val="1"/></w:numPr>',
  ),
);
body.push(
  p(
    textRun("Numbering cancelled by numId zero."),
    '<w:pStyle w:val="NumberedHeading"/><w:numPr><w:numId w:val="0"/></w:numPr>',
  ),
);
body.push(p(`<w:r><w:br w:type="page"/></w:r>`));
const rows = [
  `<w:tr><w:trPr><w:tblHeader/></w:trPr>${cell("Area", '<w:gridSpan w:val="2"/><w:shd w:fill="DCE5FF"/>')}${cell("Details", '<w:shd w:fill="DCE5FF"/>')}</w:tr>`,
  `<w:tr>${cell("Merged group", '<w:gridSpan w:val="2"/>')}${cell("First row with readable content", '<w:vMerge w:val="restart"/>')}</w:tr>`,
  `<w:tr>${cell("", '<w:gridSpan w:val="2"/>')}${cell("Second row continues the merged first cell", "<w:vMerge/>")}</w:tr>`,
  `<w:tr><w:trPr><w:ins w:id="7" w:author="Aiko" w:date="2025-02-01T08:00:00Z"/></w:trPr>${cell("New row")}${cell("Inserted table row", '<w:gridSpan w:val="2"/>')}</w:tr>`,
  `<w:tr><w:trPr><w:del w:id="8" w:author="Morgan" w:date="2025-02-02T08:00:00Z"/></w:trPr>${cell("Removed row")}${cell("Deleted table row", '<w:gridSpan w:val="2"/>')}</w:tr>`,
];
for (let i = 1; i <= 28; i++)
  rows.push(
    `<w:tr>${cell(`Row ${i}`)}${cell(`A deliberately repeated table entry to make the table continue onto additional pages (${i}).`, '<w:gridSpan w:val="2"/>')}</w:tr>`,
  );
body.push(
  `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="5000" w:type="pct"/></w:tblPr><w:tblGrid><w:gridCol w:w="1200"/><w:gridCol w:w="1200"/><w:gridCol w:w="6600"/></w:tblGrid>${rows.join("")}</w:tbl>`,
);
body.push(paragraph("Table continuation", "Heading1"));
for (let i = 1; i <= 9; i++) {
  body.push(
    paragraph(
      `Portrait content paragraph ${i}. This section uses A4 portrait geometry before the landscape section break.`,
      "Normal",
    ),
  );
}
body.push(
  p(
    textRun("Portrait section ends here."),
    '<w:sectPr><w:type w:val="nextPage"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="500" w:footer="500"/><w:docGrid w:type="lines" w:linePitch="360"/><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/></w:sectPr>',
  ),
);
for (let i = 1; i <= 8; i++) {
  body.push(
    paragraph(
      `Landscape paragraph ${i}: wide page geometry, Japanese text 日本語, and page-field output should remain readable after this section break.`,
      "Normal",
    ),
  );
}
const finalSect = `<w:sectPr><w:type w:val="nextPage"/><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="900" w:right="900" w:bottom="900" w:left="900" w:header="420" w:footer="420"/><w:docGrid w:type="linesAndChars" w:linePitch="360"/><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/></w:sectPr>`;

const comments = xml(
  `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:comment w:id="0" w:author="Aiko" w:initials="AK" w:date="2025-02-01T09:15:00Z"><w:p w14:paraId="A1000001">${textRun("Please review the marked explanation.")}</w:p></w:comment><w:comment w:id="1" w:author="Morgan" w:initials="MG" w:date="2025-02-02T10:00:00Z"><w:p w14:paraId="A1000002">${textRun("Agreed — the wording is clear.")}</w:p></w:comment></w:comments>`,
);
const commentExtended = xml(
  `<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w15:commentEx w15:paraId="A1000001" w15:done="0"/><w15:commentEx w15:paraId="A1000002" w15:paraIdParent="A1000001" w15:done="1"/></w15:commentsEx>`,
);
const numbering = xml(
  `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1.%2)"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Arial"/></w:rPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`,
);
const styles = xml(
  `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:after="120"/></w:pPr></w:pPrDefault><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="20"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:spacing w:before="100" w:after="120" w:line="300"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="BodyText"><w:name w:val="Body Text"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="180" w:line="276" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="36"/><w:color w:val="263C76"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:rPr><w:i/><w:color w:val="68738A"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="300" w:after="100"/></w:pPr><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="NumberedHeading"><w:basedOn w:val="Heading1"/><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr></w:style><w:style w:type="table" w:default="1" w:styleId="TableGrid"><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:color="auto"/><w:left w:val="single" w:sz="4" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:color="auto"/><w:right w:val="single" w:sz="4" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:color="auto"/></w:tblBorders><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tcPr><w:shd w:fill="F6F8FC"/></w:tcPr></w:style></w:styles>`,
);
const header = xml(
  `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr>${textRun("orim • Document preview")}</w:p></w:hdr>`,
);
const footer = xml(
  `<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:pPr><w:jc w:val="right"/></w:pPr>${textRun("Page ")}${f("PAGE")}${textRun(" of ")}<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>3</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>`,
);
const settings = xml(
  `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:evenAndOddHeaders/><w:defaultTabStop w:val="720"/></w:settings>`,
);
const theme = xml(
  `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme><a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface="Yu Gothic"/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/><a:ea typeface="Noto Sans JP"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`,
);
const document = xml(
  `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>${body.join("")}${finalSect}</w:body></w:document>`,
);
const relationships = xml(
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/><Relationship Id="rIdCommentsExtended" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/><Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rIdFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/><Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/sample.png"/><Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/orim" TargetMode="External"/></Relationships>`,
);
const contentTypes = xml(
  `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/><Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>`,
);
const rootRels = xml(
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
);
const wordParts = {
  "[Content_Types].xml": contentTypes,
  "_rels/.rels": rootRels,
  "word/document.xml": document,
  "word/_rels/document.xml.rels": relationships,
  "word/styles.xml": styles,
  "word/numbering.xml": numbering,
  "word/comments.xml": comments,
  "word/commentsExtended.xml": commentExtended,
  "word/settings.xml": settings,
  "word/header1.xml": header,
  "word/footer1.xml": footer,
  "word/theme/theme1.xml": theme,
  "word/media/sample.png": png(),
};
mkdirSync("fixtures", { recursive: true });
writeFileSync(
  "fixtures/sample.docx",
  zipSync(
    Object.fromEntries(
      Object.entries(wordParts).map(([path, value]) => [
        path,
        typeof value === "string" ? strToU8(value) : value,
      ]),
    ),
    { level: 6 },
  ),
);
