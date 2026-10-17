import { describe, expect, it } from "vite-plus/test";
import { formatCellValue } from "./format";

const number = (value: number) => ({ type: "number" as const, value });
const string = (value: string) => ({ type: "string" as const, value });

describe("Excel number formats", () => {
  it("formats General values", () => {
    expect(formatCellValue(number(0.1 + 0.2), "General", false).text).toBe("0.3");
    expect(formatCellValue(number(1 / 3), "General", false).text).toBe("0.333333333");
    expect(formatCellValue(number(-1 / 3), "General", false).text).toBe("-0.333333333");
    expect(formatCellValue(number(12345678901), "General", false).text).toBe("12345678901");
    expect(formatCellValue(number(1234567.891), "General", false).text).toBe("1234567.891");
    expect(formatCellValue(number(123456789012), "General", false).text).toBe("1.23457E+11");
    expect(formatCellValue(number(0.0001), "General", false).text).toBe("0.0001");
    expect(formatCellValue(number(0.00001), "General", false).text).toBe("1E-05");
    expect(formatCellValue(number(0.00012345678), "General", false).text).toBe("0.000123457");
    expect(formatCellValue(number(0.000012345), "General", false).text).toBe("1.2345E-05");
  });

  it("handles empty, text, boolean, and error values", () => {
    expect(formatCellValue({ type: "empty" }, "General", false).text).toBe("");
    expect(formatCellValue(string("hello"), "General", false).text).toBe("hello");
    expect(formatCellValue(string("hello"), '0;0;0;"value: "@', false).text).toBe("value: hello");
    expect(formatCellValue(string("hello"), '"text: "@', false).text).toBe("text: hello");
    expect(formatCellValue(string("hello"), "0.00", false).text).toBe("hello");
    expect(formatCellValue({ type: "boolean", value: true }, "General", false).text).toBe("TRUE");
    expect(formatCellValue({ type: "boolean", value: false }, "General", false).text).toBe("FALSE");
    expect(formatCellValue({ type: "error", value: "#DIV/0!" }, "General", false).text).toBe(
      "#DIV/0!",
    );
  });

  it("selects positive, negative, zero, and conditional sections", () => {
    expect(formatCellValue(number(-12.5), "0.0", false).text).toBe("-12.5");
    expect(formatCellValue(number(-12.5), "0.0;(0.0)", false).text).toBe("(12.5)");
    expect(formatCellValue(number(0), '0;-0;"zero"', false).text).toBe("zero");
    expect(formatCellValue(number(125), "[>=100]0.0;[<100]0.00", false).text).toBe("125.0");
    expect(formatCellValue(number(25), "[>=100]0.0;[<100]0.00", false).text).toBe("25.00");
    expect(formatCellValue(number(-25), "[>=100]0.0;[Red]0.00;General", false).text).toBe("-25.00");
  });

  it("formats placeholders, rounding, grouping, percent, and literals", () => {
    expect(formatCellValue(number(1234567.89), "#,##0", false).text).toBe("1,234,568");
    expect(formatCellValue(number(0.1234), "0.0%", false).text).toBe("12.3%");
    expect(formatCellValue(number(1234567), "#,##0,", false).text).toBe("1,235");
    expect(formatCellValue(number(-1234.5), "$#,##0.00", false).text).toBe("-$1,234.50");
    expect(formatCellValue(number(-5), "#,##0.00;[Red](#,##0.00)", false)).toEqual({
      text: "(5.00)",
      color: "#ff0000",
    });
    expect(formatCellValue(number(1.005), "0.00", false).text).toBe("1.01");
    expect(formatCellValue(number(2.5), "0", false).text).toBe("3");
    expect(formatCellValue(number(-2.5), "0", false).text).toBe("-3");
    expect(formatCellValue(number(1.2), "0.00", false).text).toBe("1.20");
    expect(formatCellValue(number(1.2), "0.??", false).text).toBe("1.2 ");
    expect(formatCellValue(number(1234567), "000-0000", false).text).toBe("123-4567");
    expect(formatCellValue(number(12345), "0.00E+00", false).text).toBe("1.23E+04");
    expect(formatCellValue(number(12345), "##0.0E+0", false).text).toBe("12.3E+3");
  });

  it("formats fractions with denominator limits", () => {
    expect(formatCellValue(number(1.5), "# ?/?", false).text).toBe("1 1/2");
    expect(formatCellValue(number(0.333), "# ?/?", false).text).toBe(" 1/3");
    expect(formatCellValue(number(0.5), "?/8", false).text).toBe("4/8");
  });

  it("handles format literals and bracket directives", () => {
    expect(formatCellValue(number(12), '"USD "0', false).text).toBe("USD 12");
    expect(formatCellValue(number(12), "\\$0", false).text).toBe("$12");
    expect(formatCellValue(number(12), "_x0", false).text).toBe(" 12");
    expect(formatCellValue(number(12), "*x0", false).text).toBe("12");
    expect(formatCellValue(number(12), "[$€-407]0", false).text).toBe("€12");
    expect(formatCellValue(number(12), "[$-411]0", false).text).toBe("12");
    expect(formatCellValue(number(12), "[Blue]0", false)).toEqual({
      text: "12",
      color: "#0000ff",
    });
    expect(formatCellValue(number(12), "[Color10]0", false)).toEqual({
      text: "12",
      color: "#008000",
    });
    expect(formatCellValue(number(12), "unparseable", false).text).toBe("12");
  });

  it("formats calendar dates and times in both date systems", () => {
    expect(formatCellValue(number(45000), "yyyy-mm-dd", false).text).toBe("2023-03-15");
    expect(formatCellValue(number(45000.5), "m/d/yyyy h:mm", false).text).toBe("3/15/2023 12:00");
    expect(formatCellValue(number(0.75), "h:mm AM/PM", false).text).toBe("6:00 PM");
    expect(formatCellValue(number(0.5), 'h "A/P"', false).text).toBe("12 A/P");
    expect(formatCellValue(number(45000), 'yyyy"年"m"月"d"日"', false).text).toBe("2023年3月15日");
    expect(formatCellValue(number(45000), "dddd", false).text).toBe("Wednesday");
    expect(formatCellValue(number(45000), "aaa", false).text).toBe("水");
    expect(formatCellValue(number(60), "yyyy-mm-dd", false).text).toBe("1900-02-29");
    expect(formatCellValue(number(1), "yyyy-mm-dd", true).text).toBe("1904-01-02");
    expect(formatCellValue(number(1.5), "[h]:mm:ss", false).text).toBe("36:00:00");
    expect(formatCellValue(number(-1), "yyyy-mm-dd", false).text).toBe("########");
    expect(formatCellValue(number(2958466), "yyyy-mm-dd", false).text).toBe("########");
  });
});
