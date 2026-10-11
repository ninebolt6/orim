import { describe, expect, it } from "vite-plus/test";
import { border } from "./borders";

function parseBorder(value: string): Element {
  return new DOMParser().parseFromString(
    `<w:bottom xmlns:w="urn:test" ${value}/>`,
    "application/xml",
  ).documentElement!;
}

describe("shared border conversion", () => {
  it("uses currentColor for auto, none for nil, and maps double borders", () => {
    expect(border(parseBorder('w:val="single" w:color="auto"'))).toBe("0.5pt solid currentColor");
    expect(border(parseBorder('w:val="nil"'))).toBe("none");
    expect(border(parseBorder('w:val="double" w:sz="8" w:color="123456"'))).toBe(
      "1pt double #123456",
    );
  });
});
