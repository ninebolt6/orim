import { describe, expect, it } from "vite-plus/test";
import { themeFonts } from "./theme";

const theme = (fontScheme: string) =>
  new DOMParser().parseFromString(
    `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme>${fontScheme}</a:fontScheme></a:themeElements></a:theme>`,
    "application/xml",
  );

describe("themeFonts", () => {
  it("uses a Jpan script font when the East Asian typeface is empty", () => {
    const doc = theme(`
      <a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Gothic"/><a:font script="Hang" typeface="Malgun Gothic"/><a:font script="Hans" typeface="DengXian"/></a:majorFont>
      <a:minorFont><a:latin typeface="Aptos"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Noto Sans JP"/><a:font script="Hang" typeface="Malgun Gothic"/><a:font script="Hans" typeface="DengXian"/></a:minorFont>
    `);

    expect(themeFonts(doc)).toMatchObject({
      major: "Aptos Display",
      minor: "Aptos",
      majorEastAsia: "Yu Gothic",
      minorEastAsia: "Noto Sans JP",
    });
  });

  it("keeps an explicit East Asian typeface ahead of script fonts", () => {
    const doc = theme(
      `<a:minorFont><a:latin typeface="Aptos"/><a:ea typeface="MS Mincho"/><a:cs typeface=""/><a:font script="Jpan" typeface="Noto Sans JP"/></a:minorFont>`,
    );

    expect(themeFonts(doc).minorEastAsia).toBe("MS Mincho");
  });
});
