import { mkdirSync, writeFileSync } from "node:fs";
import { deflateSync } from "node:zlib";
import { strToU8, zipSync } from "fflate";

const P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const ns = `xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}"`;
const xml = (content) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${content}`;
const EMU = 12700;
const emu = (points) => Math.round(points * EMU);

const crcTable = Array.from({ length: 256 }, (_, n) => {
  let c = n;
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  return c >>> 0;
});
const crc32 = (bytes) => {
  let c = 0xffffffff;
  for (const byte of bytes) c = crcTable[(c ^ byte) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
};
const chunk = (type, data) => {
  const body = Buffer.concat([Buffer.from(type), data]);
  const length = Buffer.alloc(4);
  length.writeUInt32BE(data.length);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(body));
  return Buffer.concat([length, body, crc]);
};
/** A 64×40 diagonal gradient PNG. */
const png = () => {
  const width = 64;
  const height = 40;
  const header = Buffer.alloc(13);
  header.writeUInt32BE(width, 0);
  header.writeUInt32BE(height, 4);
  header.set([8, 2, 0, 0, 0], 8);
  const rows = Buffer.alloc((width * 3 + 1) * height);
  for (let y = 0; y < height; y++)
    for (let x = 0; x < width; x++) {
      const offset = y * (width * 3 + 1) + 1 + x * 3;
      rows.set([40 + x * 3, 120 + y * 2, 200], offset);
    }
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk("IHDR", header),
    chunk("IDAT", deflateSync(rows)),
    chunk("IEND", Buffer.alloc(0)),
  ]);
};

const xfrm = ([x, y, width, height], attributes = "") =>
  `<a:xfrm${attributes}><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(width)}" cy="${emu(height)}"/></a:xfrm>`;
const run = (text, properties = "") =>
  `<a:r><a:rPr lang="en-US"${properties}</a:rPr><a:t>${text}</a:t></a:r>`.replace(
    '<a:rPr lang="en-US"</a:rPr>',
    '<a:rPr lang="en-US"/>',
  );
const paragraph = (content, properties = "") =>
  `<a:p>${properties ? `<a:pPr${properties}` : ""}${content}</a:p>`;
const body = (paragraphs, bodyPr = "<a:bodyPr/>") =>
  `<p:txBody>${bodyPr}<a:lstStyle/>${paragraphs}</p:txBody>`;
const sp = ({
  id,
  name,
  ph = "",
  box,
  geometry = "",
  fill = "",
  line = "",
  style = "",
  text = "",
}) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr>${ph}</p:nvPr></p:nvSpPr><p:spPr>${box ? xfrm(box) : ""}${geometry}${fill}${line}</p:spPr>${style}${text}</p:sp>`;
const preset = (name, adjust = "") =>
  `<a:prstGeom prst="${name}"><a:avLst>${adjust}</a:avLst></a:prstGeom>`;
const solid = (color) => `<a:solidFill>${color}</a:solidFill>`;
const scheme = (name, modifiers = "") => `<a:schemeClr val="${name}">${modifiers}</a:schemeClr>`;
const shapeStyle = `<p:style><a:lnRef idx="2">${scheme("accent1", '<a:shade val="50000"/>')}</a:lnRef><a:fillRef idx="1">${scheme("accent1")}</a:fillRef><a:effectRef idx="0">${scheme("accent1")}</a:effectRef><a:fontRef idx="minor">${scheme("lt1")}</a:fontRef></p:style>`;
const tree = (shapes, background = "") =>
  `<p:cSld>${background}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${shapes}</p:spTree></p:cSld>`;
const rels = (items) =>
  xml(
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items
      .map(
        ([id, type, target, external]) =>
          `<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"${external ? ' TargetMode="External"' : ""}/>`,
      )
      .join("")}</Relationships>`,
  );

const theme = xml(`<a:theme xmlns:a="${A}" name="orim"><a:themeElements>
<a:clrScheme name="orim"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F2A44"/></a:dk2><a:lt2><a:srgbClr val="EEF2F7"/></a:lt2><a:accent1><a:srgbClr val="2F6FDB"/></a:accent1><a:accent2><a:srgbClr val="E0563B"/></a:accent2><a:accent3><a:srgbClr val="2BA36B"/></a:accent3><a:accent4><a:srgbClr val="8A5CD6"/></a:accent4><a:accent5><a:srgbClr val="1BA7C4"/></a:accent5><a:accent6><a:srgbClr val="F2A33A"/></a:accent6><a:hlink><a:srgbClr val="2F6FDB"/></a:hlink><a:folHlink><a:srgbClr val="8A5CD6"/></a:folHlink></a:clrScheme>
<a:fontScheme name="orim"><a:majorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>
<a:fmtScheme name="orim"><a:fillStyleLst>${solid(scheme("phClr"))}<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0">${scheme("phClr", '<a:tint val="50000"/>')}</a:gs><a:gs pos="100000">${scheme("phClr")}</a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>${solid(scheme("phClr", '<a:shade val="80000"/>'))}</a:fillStyleLst>
<a:lnStyleLst><a:ln w="9525">${solid(scheme("phClr"))}</a:ln><a:ln w="19050">${solid(scheme("phClr"))}</a:ln><a:ln w="38100">${solid(scheme("phClr"))}</a:ln></a:lnStyleLst>
<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
<a:bgFillStyleLst>${solid(scheme("phClr"))}<a:gradFill><a:gsLst><a:gs pos="0">${scheme("phClr")}</a:gs><a:gs pos="100000">${scheme("phClr", '<a:shade val="90000"/>')}</a:gs></a:gsLst><a:lin ang="5400000"/></a:gradFill>${solid(scheme("phClr"))}</a:bgFillStyleLst></a:fmtScheme>
</a:themeElements></a:theme>`);

const level = (index, size, bullet) =>
  `<a:lvl${index}pPr marL="${emu(18 + (index - 1) * 27)}" indent="${emu(-18)}">${bullet}<a:defRPr sz="${size}"/></a:lvl${index}pPr>`;
const master = xml(
  `<p:sldMaster ${ns}>${tree(
    sp({
      id: 2,
      name: "Title Placeholder",
      ph: '<p:ph type="title"/>',
      box: [48, 24, 864, 80],
      text: body(paragraph(""), '<a:bodyPr anchor="ctr"/>'),
    }) +
      sp({
        id: 3,
        name: "Text Placeholder",
        ph: '<p:ph type="body" idx="1"/>',
        box: [48, 120, 864, 360],
        text: body(paragraph(""), "<a:bodyPr/>"),
      }) +
      sp({
        id: 4,
        name: "Slide Number",
        ph: '<p:ph type="sldNum" sz="quarter" idx="12"/>',
        box: [800, 500, 112, 28],
        text: body(
          paragraph(
            `<a:fld id="{0D2B8E5C-0000-4000-8000-000000000001}" type="slidenum"><a:rPr lang="en-US"/><a:t>‹#›</a:t></a:fld>`,
            ' algn="r"/>',
          ),
        ),
      }) +
      sp({
        id: 5,
        name: "Accent Bar",
        box: [0, 532, 960, 8],
        geometry: preset("rect"),
        fill: solid(scheme("accent1")),
      }),
    `<p:bg><p:bgRef idx="1001">${scheme("bg1")}</p:bgRef></p:bg>`,
  )}<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/><p:sldLayoutId id="2147483650" r:id="rId2"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle><a:lvl1pPr algn="l"><a:defRPr sz="3600" b="1">${solid(scheme("tx2"))}<a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr marL="${emu(18)}" indent="${emu(-18)}"><a:buFont typeface="Arial"/><a:buChar char="•"/><a:defRPr sz="2400">${solid(scheme("tx1"))}</a:defRPr></a:lvl1pPr>${level(2, 2000, '<a:buFont typeface="Arial"/><a:buChar char="–"/>')}${level(3, 1800, '<a:buFont typeface="Arial"/><a:buChar char="•"/>')}</p:bodyStyle><p:otherStyle><a:lvl1pPr><a:defRPr sz="1400">${solid(scheme("tx1", '<a:lumMod val="65000"/><a:lumOff val="35000"/>'))}</a:defRPr></a:lvl1pPr></p:otherStyle></p:txStyles></p:sldMaster>`,
);
const titleLayout = xml(
  `<p:sldLayout ${ns} type="title">${tree(
    sp({
      id: 2,
      name: "Title",
      ph: '<p:ph type="ctrTitle"/>',
      box: [96, 170, 768, 110],
      text: body(paragraph(""), '<a:bodyPr anchor="b"/>'),
    }) +
      sp({
        id: 3,
        name: "Subtitle",
        ph: '<p:ph type="subTitle" idx="1"/>',
        box: [96, 296, 768, 60],
        text: `<p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr marL="0" indent="0" algn="ctr"><a:buNone/><a:defRPr sz="2000">${solid(scheme("tx1", '<a:tint val="75000"/>'))}</a:defRPr></a:lvl1pPr></a:lstStyle><a:p/></p:txBody>`,
      }),
    `<p:bg><p:bgPr><a:gradFill><a:gsLst><a:gs pos="0">${scheme("bg1")}</a:gs><a:gs pos="100000">${scheme("bg2")}</a:gs></a:gsLst><a:lin ang="5400000"/></a:gradFill><a:effectLst/></p:bgPr></p:bg>`,
  )}<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
);
const contentLayout = xml(
  `<p:sldLayout ${ns} type="obj">${tree(
    sp({ id: 2, name: "Title", ph: '<p:ph type="title"/>' }) +
      sp({ id: 3, name: "Content", ph: '<p:ph idx="1"/>' }) +
      sp({ id: 4, name: "Slide Number", ph: '<p:ph type="sldNum" sz="quarter" idx="12"/>' }),
  )}<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
);

const slide = (shapes, attributes = "") =>
  xml(
    `<p:sld ${ns}${attributes}>${tree(shapes)}<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`,
  );
const title = (text) =>
  sp({ id: 2, name: "Title", ph: '<p:ph type="title"/>', text: body(paragraph(run(text))) });
const number = sp({
  id: 9,
  name: "Slide Number",
  ph: '<p:ph type="sldNum" sz="quarter" idx="12"/>',
  text: body(
    paragraph(
      `<a:fld id="{0D2B8E5C-0000-4000-8000-000000000001}" type="slidenum"><a:rPr lang="en-US"/><a:t>‹#›</a:t></a:fld>`,
    ),
  ),
});

const slides = [
  [
    slide(
      sp({
        id: 2,
        name: "Title",
        ph: '<p:ph type="ctrTitle"/>',
        text: body(paragraph(run("orim presentation sample"))),
      }) +
        sp({
          id: 3,
          name: "Subtitle",
          ph: '<p:ph type="subTitle" idx="1"/>',
          text: body(
            paragraph(
              run("Rendered in the browser from PresentationML") +
                run(" · ") +
                run("orim on GitHub", '><a:hlinkClick r:id="rLink"/>'),
            ),
          ),
        }),
    ),
    1,
    [["rLink", "hyperlink", "https://github.com/ninebolt6/orim", true]],
  ],
  [
    slide(
      title("Text and lists") +
        sp({
          id: 3,
          name: "Content",
          ph: '<p:ph idx="1"/>',
          text: body(
            paragraph(run("Placeholders inherit position and style from the layout and master")) +
              paragraph(run("Second-level bullets use the master body style"), ' lvl="1"/>') +
              paragraph(
                run("Runs mix ") +
                  run("bold", ' b="1">') +
                  run(", ") +
                  run("italic", ' i="1">') +
                  run(", ") +
                  run("underline", ' u="sng">') +
                  run(" and ") +
                  run("color", `>${solid(scheme("accent2"))}`),
                ' lvl="1"/>',
              ) +
              paragraph(
                run("Automatic numbering"),
                ' lvl="2"><a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabicPeriod"/></a:pPr>',
              ) +
              paragraph(
                run("Continues the sequence"),
                ' lvl="2"><a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabicPeriod"/></a:pPr>',
              ) +
              paragraph(run("日本語のテキストも表示できます")) +
              paragraph(run("Jump to the table slide", '><a:hlinkClick r:id="rSlide"/>')),
          ),
        }) +
        number,
    ),
    2,
    [["rSlide", "slide", "slide4.xml"]],
  ],
  [
    slide(
      title("Shapes, lines and pictures") +
        sp({
          id: 3,
          name: "Rounded",
          box: [48, 130, 200, 100],
          geometry: preset("roundRect"),
          style: shapeStyle,
          text: body(paragraph(run("Theme style"), ' algn="ctr"/>'), '<a:bodyPr anchor="ctr"/>'),
        }) +
        sp({
          id: 4,
          name: "Gradient Ellipse",
          box: [290, 120, 140, 120],
          geometry: preset("ellipse"),
          fill: `<a:gradFill><a:gsLst><a:gs pos="0">${scheme("accent6")}</a:gs><a:gs pos="100000">${scheme("accent2")}</a:gs></a:gsLst><a:lin ang="2700000"/></a:gradFill>`,
          line: `<a:ln w="19050">${solid(scheme("accent2", '<a:lumMod val="75000"/>'))}</a:ln>`,
        }) +
        sp({
          id: 5,
          name: "Arrow",
          box: [470, 145, 180, 70],
          geometry: preset("rightArrow"),
          fill: solid(scheme("accent3")),
          line: "<a:ln><a:noFill/></a:ln>",
        }).replace("<a:xfrm>", '<a:xfrm rot="-1200000">') +
        sp({
          id: 6,
          name: "Star",
          box: [700, 115, 130, 130],
          geometry: preset("star5"),
          fill: solid(scheme("accent4")),
          line: "<a:ln><a:noFill/></a:ln>",
        }) +
        `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="7" name="Connector"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr>${xfrm([60, 290, 360, 60])}${preset("straightConnector1")}<a:ln w="28575">${solid(scheme("accent1"))}<a:prstDash val="dash"/><a:tailEnd type="triangle"/></a:ln></p:spPr></p:cxnSp>` +
        `<p:pic><p:nvPicPr><p:cNvPr id="8" name="Picture" descr="Gradient image"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rImage"/><a:srcRect l="10000" r="10000"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${xfrm([480, 270, 200, 150])}${preset("rect")}<a:ln w="12700">${solid(scheme("tx1"))}</a:ln></p:spPr></p:pic>` +
        `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="${emu(720)}" y="${emu(280)}"/><a:ext cx="${emu(180)}" cy="${emu(120)}"/><a:chOff x="0" y="0"/><a:chExt cx="${emu(360)}" cy="${emu(240)}"/></a:xfrm></p:grpSpPr>${sp({ id: 11, name: "Group Box", box: [0, 0, 360, 240], geometry: preset("rect"), fill: solid(scheme("accent5", '<a:lumMod val="20000"/><a:lumOff val="80000"/>')), line: `<a:ln w="12700">${solid(scheme("accent5"))}</a:ln>` })}${sp({ id: 12, name: "Group Label", box: [20, 80, 320, 80], text: body(paragraph(run("Grouped shapes", ' sz="2800" b="1">' + solid(scheme("accent5", '<a:lumMod val="50000"/>'))), ' algn="ctr"/>'), '<a:bodyPr anchor="ctr"/>') })}</p:grpSp>` +
        sp({
          id: 13,
          name: "Caption",
          box: [48, 440, 400, 40],
          text: body(
            paragraph(
              run("Rotation, gradients, dashed connectors and cropped pictures.", ' sz="1400">'),
            ),
          ),
        }) +
        number,
    ),
    2,
    [["rImage", "image", "../media/image1.png"]],
  ],
  [
    slide(
      title("Tables") +
        `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="Table"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${emu(48)}" y="${emu(130)}"/><a:ext cx="${emu(864)}" cy="${emu(160)}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId></a:tblPr><a:tblGrid>${[288, 192, 192, 192].map((width) => `<a:gridCol w="${emu(width)}"/>`).join("")}</a:tblGrid>${[
          ["Format", "Parser", "Renderer", "React"],
          ["DOCX", "✓", "✓", "DocxViewer"],
          ["XLSX", "✓", "✓", "XlsxViewer"],
          ["PPTX", "✓", "✓", "PptxViewer"],
        ]
          .map(
            (row) =>
              `<a:tr h="${emu(40)}">${row
                .map(
                  (cell) =>
                    `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1800"/><a:t>${cell}</a:t></a:r></a:p></a:txBody><a:tcPr anchor="ctr"/></a:tc>`,
                )
                .join("")}</a:tr>`,
          )
          .join("")}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>` +
        number,
    ),
    2,
    [],
  ],
  [slide(title("Hidden slide") + number, ' show="0"'), 2, []],
];

const files = {
  "[Content_Types].xml": xml(
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>${[
      1, 2,
    ]
      .map(
        (index) =>
          `<Override PartName="/ppt/slideLayouts/slideLayout${index}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>`,
      )
      .join("")}${slides
      .map(
        (_, index) =>
          `<Override PartName="/ppt/slides/slide${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`,
      )
      .join(
        "",
      )}<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>`,
  ),
  "_rels/.rels": rels([["rId1", "officeDocument", "ppt/presentation.xml"]]),
  "ppt/presentation.xml": xml(
    `<p:presentation ${ns}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>${slides
      .map((_, index) => `<p:sldId id="${256 + index}" r:id="rId${index + 3}"/>`)
      .join(
        "",
      )}</p:sldIdLst><p:sldSz cx="${emu(960)}" cy="${emu(540)}"/><p:notesSz cx="6858000" cy="9144000"/><p:defaultTextStyle><a:lvl1pPr><a:defRPr sz="1800">${solid(scheme("tx1"))}<a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr></p:defaultTextStyle></p:presentation>`,
  ),
  "ppt/_rels/presentation.xml.rels": rels([
    ["rId1", "slideMaster", "slideMasters/slideMaster1.xml"],
    ["rId2", "theme", "theme/theme1.xml"],
    ...slides.map((_, index) => [`rId${index + 3}`, "slide", `slides/slide${index + 1}.xml`]),
  ]),
  "ppt/theme/theme1.xml": theme,
  "ppt/slideMasters/slideMaster1.xml": master,
  "ppt/slideMasters/_rels/slideMaster1.xml.rels": rels([
    ["rId1", "slideLayout", "../slideLayouts/slideLayout1.xml"],
    ["rId2", "slideLayout", "../slideLayouts/slideLayout2.xml"],
    ["rId3", "theme", "../theme/theme1.xml"],
  ]),
  "ppt/slideLayouts/slideLayout1.xml": titleLayout,
  "ppt/slideLayouts/slideLayout2.xml": contentLayout,
  ...Object.fromEntries(
    [1, 2].map((index) => [
      `ppt/slideLayouts/_rels/slideLayout${index}.xml.rels`,
      rels([["rId1", "slideMaster", "../slideMasters/slideMaster1.xml"]]),
    ]),
  ),
  ...Object.fromEntries(
    slides.flatMap(([content, layout, links], index) => [
      [`ppt/slides/slide${index + 1}.xml`, content],
      [
        `ppt/slides/_rels/slide${index + 1}.xml.rels`,
        rels([["rLayout", "slideLayout", `../slideLayouts/slideLayout${layout}.xml`], ...links]),
      ],
    ]),
  ),
};

const archive = Object.fromEntries(
  Object.entries(files).map(([path, content]) => [path, strToU8(content)]),
);
archive["ppt/media/image1.png"] = new Uint8Array(png());
mkdirSync(new URL("../fixtures/", import.meta.url), { recursive: true });
writeFileSync(
  new URL("../fixtures/sample.pptx", import.meta.url),
  zipSync(archive, { level: 9, mtime: new Date("2000-01-01T00:00:00.000Z") }),
);
