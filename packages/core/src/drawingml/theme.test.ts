import { describe, expect, it } from "vite-plus/test";
import { eastAsianScript, themeEastAsiaFont, themeFonts } from "./theme";

const theme = (fontScheme: string) =>
  new DOMParser().parseFromString(
    `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme>${fontScheme}</a:fontScheme></a:themeElements></a:theme>`,
    "application/xml",
  );

describe("themeFonts", () => {
  it("reads East Asian script font supplements", () => {
    const doc = theme(`
      <a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Gothic"/><a:font script="Hang" typeface="Malgun Gothic"/><a:font script="Hans" typeface="DengXian"/></a:majorFont>
      <a:minorFont><a:latin typeface="Aptos"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Noto Sans JP"/><a:font script="Hang" typeface="Malgun Gothic"/><a:font script="Hans" typeface="DengXian"/></a:minorFont>
    `);

    const fonts = themeFonts(doc);
    expect(fonts).toMatchObject({ major: "Aptos Display", minor: "Aptos" });
    expect(fonts.majorEastAsia).toBeUndefined();
    expect(fonts.minorEastAsia).toBeUndefined();
    expect(fonts.majorEastAsiaScripts).toEqual({
      Jpan: "Yu Gothic",
      Hang: "Malgun Gothic",
      Hans: "DengXian",
    });
    expect(fonts.minorEastAsiaScripts).toEqual({
      Jpan: "Noto Sans JP",
      Hang: "Malgun Gothic",
      Hans: "DengXian",
    });
  });

  it("keeps an explicit East Asian typeface ahead of script fonts", () => {
    const doc = theme(
      `<a:minorFont><a:latin typeface="Aptos"/><a:ea typeface="MS Mincho"/><a:cs typeface=""/><a:font script="Jpan" typeface="Noto Sans JP"/></a:minorFont>`,
    );

    const fonts = themeFonts(doc);
    expect(themeEastAsiaFont(fonts, false, "ja-JP", "日本語のテキスト")).toBe("MS Mincho");
  });

  it("infers the run script from text or language", () => {
    expect(eastAsianScript(undefined, "日本語のテキスト")).toBe("Jpan");
    expect(eastAsianScript(undefined, "한국어")).toBe("Hang");
    expect(eastAsianScript("zh-CN", "汉字")).toBe("Hans");
    expect(eastAsianScript("zh-Hant-TW", "漢字")).toBe("Hant");
    expect(eastAsianScript("en-US", "abc")).toBeUndefined();
  });

  it("selects the supplement matching the run script", () => {
    const doc = theme(
      `<a:minorFont><a:latin typeface="Aptos"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Noto Sans JP"/><a:font script="Hang" typeface="Malgun Gothic"/><a:font script="Hans" typeface="DengXian"/><a:font script="Hant" typeface="PMingLiU"/></a:minorFont>`,
    );
    const fonts = themeFonts(doc);

    expect(themeEastAsiaFont(fonts, false, undefined, "日本語のテキスト")).toBe("Noto Sans JP");
    expect(themeEastAsiaFont(fonts, false, undefined, "한국어")).toBe("Malgun Gothic");
    expect(themeEastAsiaFont(fonts, false, "zh-CN", "汉字")).toBe("DengXian");
    expect(themeEastAsiaFont(fonts, false, "zh-TW", "漢字")).toBe("PMingLiU");
  });
});
