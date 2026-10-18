import { mkdirSync, writeFileSync } from "node:fs";
import { strToU8, zipSync } from "fflate";

const xml = (content) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${content}`;
const escape = (value) =>
  String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
const cell = (reference, value, style = 0, type = "n") =>
  `<c r="${reference}"${type === "n" ? "" : ` t="${type}"`}${style ? ` s="${style}"` : ""}>${value}</c>`;
const inline = (reference, value, style = 0) =>
  cell(reference, `<is><t xml:space="preserve">${escape(value)}</t></is>`, style, "inlineStr");
const row = (index, cells, attributes = "") =>
  `<row r="${index}"${attributes ? ` ${attributes}` : ""}>${cells.join("")}</row>`;
const shared = [
  `<si><t>Revenue</t></si>`,
  `<si><r><rPr><b/></rPr><t>Rich </t></r><r><rPr><color rgb="FFFF0000"/></rPr><t>text</t></r><rPh sb="0" eb="4"><t>フリガナ</t></rPh></si>`,
];

const summary = xml(
  `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetViews><sheetView showGridLines="1" showRowColHeaders="1" workbookViewId="0"><pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols><col min="1" max="1" width="14" customWidth="1"/><col min="2" max="6" width="12" customWidth="1"/><col min="7" max="7" width="12" hidden="1" customWidth="1"/></cols><sheetFormatPr defaultRowHeight="15"/><sheetData>${[
    row(1, [inline("A1", "orim spreadsheet sample", 1)], 'ht="28" customHeight="1"'),
    row(
      3,
      ["A3", "B3", "C3", "D3", "E3", "F3"].map((ref, index) =>
        inline(ref, ["Item", "Value", "Type", "Status", "Notes", "Link"][index], 2),
      ),
    ),
    row(4, [
      cell("A4", "<v>0</v>", 0, "s"),
      cell("B4", "<v>1234567</v>", 3),
      inline("C4", "Number"),
      inline("D4", "Ready"),
      cell("E4", `<f>SUM(B4:B8)</f><v>1250</v>`, 3),
      inline("F4", "orim on GitHub"),
      inline("G4", "HIDDEN COLUMN"),
    ]),
    row(5, [
      inline("A5", "Rich text"),
      cell("B5", "<v>0.1234</v>", 4),
      cell("C5", "<v>1</v>", 0, "b"),
      inline("D5", "日本語テキスト"),
      inline("E5", "_x000D_escaped text"),
      inline("F5", "Unsafe link"),
    ]),
    row(6, [
      cell("A6", "<v>1</v>", 0, "s"),
      cell("B6", "<v>-1234.5</v>", 5),
      cell("C6", "<v>#DIV/0!</v>", 0, "e"),
      inline("D6", "Wrapped content"),
      inline("E6", "Text"),
      inline("F6", "javascript link"),
    ]),
    row(7, [
      inline("A7", "Date"),
      cell("B7", "<v>45000</v>", 6),
      inline("C7", "Date format"),
      inline("D7", "Center", 7),
      inline("E7", "A long unwrapped value spills into the empty cells to its right", 8),
    ]),
    row(
      8,
      [
        inline(
          "A8",
          "This long wrapped cell stays within its column and makes the row taller.",
          14,
        ),
        cell("B8", "<v>0.75</v>", 9),
        inline("C8", "Time"),
        inline("D8", "Right aligned", 10),
        inline("E8", "Other"),
      ],
      'ht="45" customHeight="1"',
    ),
    row(9, [
      inline("A9", "Negative"),
      cell("B9", "<v>-5</v>", 11),
      inline("C9", "Red format"),
      inline("D9", "Other"),
      inline("E9", "Other"),
    ]),
    row(10, [inline("A10", "HIDDEN ROW")], 'hidden="1"'),
  ].join(
    "",
  )}</sheetData><mergeCells count="1"><mergeCell ref="A1:F1"/></mergeCells><hyperlinks><hyperlink ref="F4" r:id="safe"/><hyperlink ref="F6" r:id="unsafe"/></hyperlinks></worksheet>`,
);
const dataRows = [
  row(
    1,
    Array.from({ length: 8 }, (_, index) =>
      inline(`${String.fromCharCode(65 + index)}1`, `Column ${index + 1}`, 2),
    ),
  ),
  ...Array.from({ length: 200 }, (_, rowIndex) =>
    row(
      rowIndex + 2,
      Array.from({ length: 8 }, (_, columnIndex) =>
        cell(
          `${String.fromCharCode(65 + columnIndex)}${rowIndex + 2}`,
          `<v>${rowIndex * 8 + columnIndex}</v>`,
        ),
      ),
    ),
  ),
  row(205, [inline("B205", "Merged data block")]),
  row(206, [inline("B206", "_x000D_escaped value")]),
];
const data = xml(
  `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView showGridLines="1"><pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/></sheetView></sheetViews><sheetData>${dataRows.join("")}</sheetData><mergeCells count="1"><mergeCell ref="B205:D206"/></mergeCells></worksheet>`,
);
const hidden = xml(
  `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${row(1, [inline("A1", "SECRET SHEET")])}</sheetData></worksheet>`,
);
const noGrid = xml(
  `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView showGridLines="0"/></sheetViews><sheetData>${[
    row(1, [inline("A1", "Boxed table", 15), inline("B1", "Status", 16)]),
    row(2, [inline("A2", "Outer border", 17), inline("B2", "Dashed inside", 18)]),
  ].join("")}</sheetData></worksheet>`,
);
const workbook = xml(
  `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView activeTab="0"/></bookViews><sheets><sheet name="Summary" sheetId="1" r:id="sheet1"/><sheet name="Data" sheetId="2" r:id="sheet2"/><sheet name="Hidden" sheetId="3" state="hidden" r:id="sheet3"/><sheet name="No grid" sheetId="4" r:id="sheet4"/></sheets></workbook>`,
);
const workbookRels = xml(
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="sheet1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="../worksheets/tab-one.xml"/><Relationship Id="sheet2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="../worksheets/data.xml"/><Relationship Id="sheet3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="../worksheets/hidden.xml"/><Relationship Id="sheet4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="../worksheets/no-grid.xml"/><Relationship Id="shared" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="../shared/strings.xml"/><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="../format/cell-styles.xml"/><Relationship Id="theme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/colors.xml"/></Relationships>`,
);
const styleSheet = xml(
  `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="5"><numFmt numFmtId="164" formatCode="0.0%"/><numFmt numFmtId="165" formatCode="$#,##0.00"/><numFmt numFmtId="166" formatCode="yyyy-mm-dd"/><numFmt numFmtId="167" formatCode="#,##0.00;[Red](#,##0.00)"/><numFmt numFmtId="168" formatCode="h:mm AM/PM"/></numFmts><fonts count="3"><font><sz val="11"/><name val="Arial"/></font><font><b/><sz val="16"/><name val="Arial"/></font><font><b/><color rgb="FFFFFFFF"/><name val="Arial"/></font></fonts><fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor theme="4" tint="0.4"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="5"><border/><border><bottom style="thin"><color rgb="FF808080"/></bottom></border><border><bottom style="double"><color rgb="FF000000"/></bottom></border><border><left style="thick"><color rgb="FF000000"/></left><right style="thick"><color rgb="FF000000"/></right><top style="thick"><color rgb="FF000000"/></top><bottom style="thick"><color rgb="FF000000"/></bottom></border><border><bottom style="dashed"><color rgb="FF666666"/></bottom></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="15"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center"/></xf><xf numFmtId="3" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="2" xfId="0"/><xf numFmtId="168" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="right"/></xf><xf numFmtId="167" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="3" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="4" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`,
);
const fixtureStyles = styleSheet
  .replace('borders count="5"', 'borders count="9"')
  .replace(
    "</borders>",
    `<border><left style="thick"><color rgb="FF000000"/></left><top style="thick"><color rgb="FF000000"/></top><right style="dashed"><color rgb="FF666666"/></right><bottom style="dashed"><color rgb="FF666666"/></bottom></border><border><right style="thick"><color rgb="FF000000"/></right><top style="thick"><color rgb="FF000000"/></top><left style="dashed"><color rgb="FF666666"/></left><bottom style="dashed"><color rgb="FF666666"/></bottom></border><border><left style="thick"><color rgb="FF000000"/></left><bottom style="thick"><color rgb="FF000000"/></bottom><top style="dashed"><color rgb="FF666666"/></top><right style="dashed"><color rgb="FF666666"/></right></border><border><right style="thick"><color rgb="FF000000"/></right><bottom style="thick"><color rgb="FF000000"/></bottom><top style="dashed"><color rgb="FF666666"/></top><left style="dashed"><color rgb="FF666666"/></left></border></borders>`,
  )
  .replace('cellXfs count="15"', 'cellXfs count="19"')
  .replace(
    "</cellXfs>",
    `<xf numFmtId="0" fontId="0" fillId="0" borderId="5" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="6" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="7" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="8" xfId="0"/></cellXfs>`,
  );
const sharedStrings = xml(
  `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${shared.length}" uniqueCount="${shared.length}">${shared.join("")}</sst>`,
);
const theme = xml(
  `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme></a:themeElements></a:theme>`,
);
const rootRels = xml(
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="workbook" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="excel/book.xml"/></Relationships>`,
);
const summaryRels = xml(
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="safe" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://github.com/ninebolt6/orim" TargetMode="External"/><Relationship Id="unsafe" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="javascript:alert(1)" TargetMode="External"/></Relationships>`,
);
const overrides = [
  ["/excel/book.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"],
  [
    "/worksheets/tab-one.xml",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
  ],
  [
    "/worksheets/data.xml",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
  ],
  [
    "/worksheets/hidden.xml",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
  ],
  [
    "/worksheets/no-grid.xml",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
  ],
  [
    "/shared/strings.xml",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
  ],
  [
    "/format/cell-styles.xml",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
  ],
  ["/theme/colors.xml", "application/vnd.openxmlformats-officedocument.theme+xml"],
]
  .map(([part, type]) => `<Override PartName="${part}" ContentType="${type}"/>`)
  .join("");
const contentTypes = xml(
  `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${overrides}</Types>`,
);
const files = {
  "[Content_Types].xml": contentTypes,
  "_rels/.rels": rootRels,
  "excel/book.xml": workbook,
  "excel/_rels/book.xml.rels": workbookRels,
  "worksheets/tab-one.xml": summary,
  "worksheets/_rels/tab-one.xml.rels": summaryRels,
  "worksheets/data.xml": data,
  "worksheets/hidden.xml": hidden,
  "worksheets/no-grid.xml": noGrid,
  "shared/strings.xml": sharedStrings,
  "format/cell-styles.xml": fixtureStyles,
  "theme/colors.xml": theme,
};
mkdirSync(new URL("../fixtures/", import.meta.url), { recursive: true });
writeFileSync(
  new URL("../fixtures/sample.xlsx", import.meta.url),
  zipSync(
    Object.fromEntries(Object.entries(files).map(([path, contents]) => [path, strToU8(contents)])),
    { mtime: new Date("2000-01-01T00:00:00.000Z") },
  ),
);
