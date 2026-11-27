import { describe, expect, it } from "vite-plus/test";
import { unzipSync, zipSync } from "fflate";
import { buildPptx, defaultTheme, shape } from "../../test/pptx";
import { buildXlsx } from "../../test/xlsx";
import type { GroupElement, PictureElement, ShapeElement, TableElement } from "./model";
import { parsePptx } from "./parse";

const EMU = 12700;
const text = (value: string, properties = "") =>
  `<a:p>${properties ? `<a:pPr${properties}/>` : ""}<a:r><a:rPr lang="en-US"/><a:t>${value}</a:t></a:r></a:p>`;
const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

const master = {
  shapes:
    shape({
      id: 2,
      placeholder: '<p:ph type="title"/>',
      xfrm: [10 * EMU, 20 * EMU, 300 * EMU, 40 * EMU],
      body: "<a:p/>",
    }) +
    shape({
      id: 3,
      placeholder: '<p:ph type="body" idx="1"/>',
      xfrm: [10 * EMU, 80 * EMU, 300 * EMU, 200 * EMU],
      body: "<a:p/>",
    }) +
    shape({
      id: 4,
      name: "Master logo",
      xfrm: [0, 0, 20 * EMU, 20 * EMU],
      properties:
        '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent2"/></a:solidFill>',
    }),
  extra:
    '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>' +
    '<p:txStyles><p:titleStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="4400"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr></p:titleStyle>' +
    '<p:bodyStyle><a:lvl1pPr marL="342900" indent="-342900"><a:buChar char="•"/><a:defRPr sz="3200"/></a:lvl1pPr><a:lvl2pPr marL="742950" indent="-285750"><a:buChar char="–"/><a:defRPr sz="2800"/></a:lvl2pPr></p:bodyStyle><p:otherStyle/></p:txStyles>',
  background: '<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg2"/></p:bgRef></p:bg>',
};

const layout = {
  shapes: shape({
    id: 2,
    placeholder: '<p:ph type="body" idx="1"/>',
    xfrm: [50 * EMU, 90 * EMU, 400 * EMU, 300 * EMU],
  }),
};

describe("PPTX package parsing", () => {
  it("follows relationships and applies master/layout placeholder inheritance", async () => {
    const deck = await parsePptx(
      buildPptx({
        presentationPath: "deck/main.xml",
        size: { cx: 960 * EMU, cy: 540 * EMU },
        master,
        layout,
        slides: [
          {
            shapes:
              shape({ id: 2, placeholder: '<p:ph type="ctrTitle"/>', body: text("Title") }) +
              shape({
                id: 3,
                placeholder: '<p:ph idx="1"/>',
                body: text("First") + text("Second", ' lvl="1"'),
              }),
          },
        ],
      }),
    );
    expect(deck).toMatchObject({ width: 960, height: 540 });
    const [logo, title, body] = deck.slides[0]!.elements as ShapeElement[];
    expect(deck.slides[0]!.elements).toHaveLength(3);
    expect(logo).toMatchObject({ name: "Master logo", fill: { type: "solid", color: "#c0504d" } });
    expect(title!.transform).toMatchObject({ x: 10, y: 20, width: 300, height: 40 });
    expect(title!.text!.paragraphs[0]).toMatchObject({ align: "center" });
    expect(title!.text!.paragraphs[0]!.runs[0]).toMatchObject({
      text: "Title",
      style: { size: 44, color: "#000000", font: "Major Sans" },
    });
    expect(body!.transform).toMatchObject({ x: 50, y: 90, width: 400, height: 300 });
    expect(body!.text!.paragraphs.map((paragraph) => paragraph.bullet?.text)).toEqual(["•", "–"]);
    expect(body!.text!.paragraphs[1]).toMatchObject({ level: 1, marginLeft: 58.5, indent: -22.5 });
    expect(deck.slides[0]!.background).toEqual({ type: "solid", color: "#eeece1" });
  });

  it("falls back to an inherited bullet font when an East Asian theme reference cannot resolve", async () => {
    const deck = await parsePptx(
      buildPptx({
        defaultTextStyle:
          '<p:defaultTextStyle><a:lvl1pPr><a:buFont typeface="Wingdings"/><a:buChar char="ü"/></a:lvl1pPr></p:defaultTextStyle>',
        slides: [
          {
            shapes: shape({
              id: 2,
              xfrm: [0, 0, 400 * EMU, 100 * EMU],
              body: '<a:p><a:pPr><a:buFont typeface="+mn-ea"/></a:pPr><a:r><a:rPr lang="ja-JP"><a:ea typeface="+mn-ea"/></a:rPr><a:t>日本語</a:t></a:r></a:p>',
            }),
          },
        ],
      }),
    );

    expect(deck.slides[0]!.elements[0]).toMatchObject({
      text: {
        paragraphs: [{ bullet: { text: "✓" } }],
      },
    });
  });

  it("does not fall through a bullet font that follows the text", async () => {
    const deck = await parsePptx(
      buildPptx({
        defaultTextStyle:
          '<p:defaultTextStyle><a:defPPr><a:buFont typeface="Wingdings"/><a:buChar char="ü"/></a:defPPr><a:lvl1pPr><a:buFontTx/></a:lvl1pPr></p:defaultTextStyle>',
        slides: [
          {
            shapes: shape({
              id: 2,
              xfrm: [0, 0, 400 * EMU, 100 * EMU],
              body: '<a:p><a:pPr><a:buFont typeface="+mn-ea"/></a:pPr><a:r><a:rPr lang="ja-JP"><a:ea typeface="+mn-ea"/></a:rPr><a:t>日本語</a:t></a:r></a:p>',
            }),
          },
        ],
      }),
    );

    expect(deck.slides[0]!.elements[0]).toMatchObject({
      text: {
        paragraphs: [{ bullet: { text: "ü" } }],
      },
    });
  });

  it("resolves an empty East Asian theme typeface through matching script fonts", async () => {
    const scriptTheme = defaultTheme.replace(
      '<a:minorFont><a:latin typeface="Minor Sans"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>',
      '<a:minorFont><a:latin typeface="Minor Sans"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="MS 明朝"/><a:font script="Hang" typeface="Malgun Gothic"/><a:font script="Hans" typeface="宋体"/></a:minorFont>',
    );
    const deck = await parsePptx(
      buildPptx({
        theme: scriptTheme,
        slides: [
          {
            shapes: shape({
              id: 2,
              xfrm: [0, 0, 400 * EMU, 100 * EMU],
              body: '<a:p><a:pPr><a:buFont typeface="+mn-ea"/><a:buChar char="・"/></a:pPr><a:r><a:rPr lang="ja-JP"><a:ea typeface="+mn-ea"/></a:rPr><a:t>日本語</a:t></a:r><a:r><a:rPr lang="zh-CN"><a:ea typeface="+mn-ea"/></a:rPr><a:t>汉字文本</a:t></a:r><a:r><a:rPr><a:ea typeface="+mn-ea"/></a:rPr><a:t>한국어</a:t></a:r></a:p>',
            }),
          },
        ],
      }),
    );

    expect(deck.slides[0]!.elements[0]).toMatchObject({
      text: {
        paragraphs: [
          {
            bullet: { font: "MS 明朝" },
            runs: [
              { text: "日本語", style: { eastAsianFont: "MS 明朝" } },
              { text: "汉字文本", style: { eastAsianFont: "宋体" } },
              { text: "한국어", style: { eastAsianFont: "Malgun Gothic" } },
            ],
          },
        ],
      },
    });
  });

  it("honours hidden slides, showMasterSp and color map overrides", async () => {
    const deck = await parsePptx(
      buildPptx({
        master,
        slides: [
          {
            attributes: 'show="0" showMasterSp="0"',
            shapes: shape({
              id: 5,
              xfrm: [0, 0, EMU, EMU],
              properties: '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>',
            }),
            extra:
              '<p:clrMapOvr><a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:clrMapOvr>',
          },
        ],
      }),
    );
    const slide = deck.slides[0]!;
    expect(slide.hidden).toBe(true);
    expect(slide.elements).toHaveLength(1);
    expect(slide.elements[0]).toMatchObject({ fill: { type: "solid", color: "#ffffff" } });
  });

  it("parses pictures, groups, tables, fields and sanitized hyperlinks", async () => {
    const deck = await parsePptx(
      buildPptx({
        media: { "media/image1.png": png },
        tableStyles:
          '<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{X}"><a:tblStyle styleId="{X}" styleName="Test"><a:wholeTbl><a:tcTxStyle><a:schemeClr val="dk1"/></a:tcTxStyle><a:tcStyle><a:fill><a:solidFill><a:srgbClr val="EEEEEE"/></a:solidFill></a:fill></a:tcStyle></a:wholeTbl><a:firstRow><a:tcTxStyle b="on"/><a:tcStyle><a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle></a:firstRow></a:tblStyle></a:tblStyleLst>',
        slides: [
          {
            relationships: [
              { id: "rImage", type: "image", target: "../media/image1.png" },
              { id: "rSafe", type: "hyperlink", target: "https://example.com/", external: true },
              { id: "rBad", type: "hyperlink", target: "javascript:alert(1)", external: true },
              { id: "rSlide", type: "slide", target: "slide2.xml" },
            ],
            shapes:
              `<p:pic><p:nvPicPr><p:cNvPr id="2" name="Picture" descr="Logo"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rImage"/><a:srcRect l="10000" r="10000"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${EMU * 10}" cy="${EMU * 10}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>` +
              `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm rot="5400000" flipH="1"><a:off x="${EMU * 100}" y="0"/><a:ext cx="${EMU * 50}" cy="${EMU * 50}"/><a:chOff x="0" y="0"/><a:chExt cx="${EMU * 100}" cy="${EMU * 100}"/></a:xfrm></p:grpSpPr>${shape(
                {
                  id: 4,
                  xfrm: [0, 0, EMU * 100, EMU * 100],
                  nonVisual: "",
                  body: '<a:p><a:r><a:rPr><a:hlinkClick r:id="rSafe"/></a:rPr><a:t>safe</a:t></a:r><a:r><a:rPr><a:hlinkClick r:id="rBad"/></a:rPr><a:t>bad</a:t></a:r><a:r><a:rPr><a:hlinkClick r:id="rSlide"/></a:rPr><a:t>next</a:t></a:r><a:fld id="{1}" type="slidenum"><a:t>‹#›</a:t></a:fld></a:p>',
                },
              )}</p:grpSp>` +
              `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="5" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="${EMU * 200}"/><a:ext cx="${EMU * 200}" cy="${EMU * 40}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1"><a:tableStyleId>{X}</a:tableStyleId></a:tblPr><a:tblGrid><a:gridCol w="${EMU * 100}"/><a:gridCol w="${EMU * 100}"/></a:tblGrid><a:tr h="${EMU * 20}"><a:tc gridSpan="2"><a:txBody><a:bodyPr/><a:p><a:r><a:t>Head</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc><a:tc hMerge="1"><a:txBody><a:bodyPr/><a:p/></a:txBody><a:tcPr/></a:tc></a:tr><a:tr h="${EMU * 20}"><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>A</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>B</a:t></a:r></a:p></a:txBody><a:tcPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:tcPr></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>`,
          },
          { shapes: "" },
        ],
      }),
    );
    const [picture, group, table] = deck.slides[0]!.elements as [
      PictureElement,
      GroupElement,
      TableElement,
    ];
    expect(picture).toMatchObject({
      kind: "picture",
      part: "ppt/media/image1.png",
      description: "Logo",
      crop: { left: 0.1, right: 0.1, top: 0, bottom: 0 },
    });
    expect(deck.media["ppt/media/image1.png"]?.mime).toBe("image/png");
    expect(Array.from(deck.media["ppt/media/image1.png"]!.bytes.slice())).toEqual(Array.from(png));
    expect(group).toMatchObject({
      kind: "group",
      transform: { x: 100, rotation: 90, flipH: true },
      childBounds: { x: 0, y: 0, width: 100, height: 100 },
    });
    const runs = (group.children[0] as ShapeElement).text!.paragraphs[0]!.runs;
    expect(runs.map((run) => ("href" in run ? run.href : undefined))).toEqual([
      "https://example.com/",
      undefined,
      "#slide-2",
      undefined,
    ]);
    expect(runs.map((run) => ("text" in run ? run.text : ""))).toEqual([
      "safe",
      "bad",
      "next",
      "1",
    ]);
    expect(table.columns).toEqual([100, 100]);
    expect(table.rows[0]!.cells[0]).toMatchObject({
      columnSpan: 2,
      fill: { type: "solid", color: "#4f81bd" },
    });
    expect(table.rows[0]!.cells[0]!.text!.paragraphs[0]!.runs[0]).toMatchObject({
      style: { bold: true },
    });
    expect(table.rows[0]!.cells[1]).toMatchObject({ merged: true });
    expect(table.rows[1]!.cells[0]).toMatchObject({ fill: { type: "solid", color: "#eeeeee" } });
    expect(table.rows[1]!.cells[1]).toMatchObject({ fill: { type: "solid", color: "#ff0000" } });
  });

  it("rejects packages without a presentation part", async () => {
    await expect(parsePptx(buildXlsx({ sheets: [] }))).rejects.toThrow(/Not a PPTX package/);
    await expect(parsePptx(new Blob([buildPptx({ slides: [{}] }).slice()]))).resolves.toMatchObject(
      {
        slides: [{ number: 1, hidden: false }],
      },
    );
  });

  it("renders slide content when the layout part is missing", async () => {
    const entries = unzipSync(
      buildPptx({
        slides: [{ shapes: shape({ id: 2, xfrm: [0, 0, EMU, EMU], body: text("Solo") }) }],
      }),
    );
    delete entries["ppt/slideLayouts/slideLayout1.xml"];
    const deck = await parsePptx(zipSync(entries, { level: 0 }));
    expect(deck.slides[0]!.elements).toHaveLength(1);
    expect(deck.slides[0]!.elements[0]).toMatchObject({
      kind: "shape",
      text: { paragraphs: [{ runs: [{ text: "Solo" }] }] },
    });
  });
});
