import { attr, children, elements, first } from "../../common/xml";
import type { BorderEdge, BorderLine, BorderSet, CellStyle, FontStyle } from "../model";
import { indexedColors, parseColor } from "../colors";

const builtInFormats: Readonly<Record<number, string>> = {
  0: "General",
  1: "0",
  2: "0.00",
  3: "#,##0",
  4: "#,##0.00",
  9: "0%",
  10: "0.00%",
  11: "0.00E+00",
  12: "# ?/?",
  13: "# ??/??",
  14: "m/d/yyyy",
  15: "d-mmm-yy",
  16: "d-mmm",
  17: "mmm-yy",
  18: "h:mm AM/PM",
  19: "h:mm:ss AM/PM",
  20: "h:mm",
  21: "h:mm:ss",
  22: "m/d/yyyy h:mm",
  37: "#,##0 ;(#,##0)",
  38: "#,##0 ;[Red](#,##0)",
  39: "#,##0.00;(#,##0.00)",
  40: "#,##0.00;[Red](#,##0.00)",
  45: "mm:ss",
  46: "[h]:mm:ss",
  47: "mmss.0",
  48: "##0.0E+0",
  49: "@",
};

const enabled = (element: Element | undefined): boolean => {
  const value = attr(element, "val");
  return element !== undefined && !["0", "false", "off", "no"].includes(value?.toLowerCase() ?? "");
};

const numberAttribute = (element: Element | undefined, name: string): number | undefined => {
  const value = Number(attr(element, name));
  return Number.isFinite(value) ? value : undefined;
};

const readFont = (
  element: Element | undefined,
  theme: readonly string[],
  palette: readonly string[],
  rich = false,
): FontStyle => {
  const font = element ? children(element) : [];
  const nameElement = font.find((child) => child.localName === (rich ? "rFont" : "name"));
  const size = numberAttribute(
    font.find((child) => child.localName === "sz"),
    "val",
  );
  const name = attr(nameElement, "val");
  const underline = font.find((child) => child.localName === "u");
  const underlineValue = underline ? (attr(underline, "val") ?? "single") : undefined;
  const vertical = font.find((child) => child.localName === "vertAlign");
  const verticalValue = attr(vertical, "val");
  const colorElement = font.find((child) => child.localName === "color");
  const style: FontStyle = {
    bold: enabled(font.find((child) => child.localName === "b")),
    italic: enabled(font.find((child) => child.localName === "i")),
    strike: enabled(font.find((child) => child.localName === "strike")),
    ...(name ? { name } : {}),
    ...(size !== undefined && size > 0 ? { size } : {}),
    ...(underlineValue === "single" || underlineValue === "double"
      ? { underline: underlineValue }
      : {}),
    ...(verticalValue === "superscript" || verticalValue === "subscript"
      ? { vertAlign: verticalValue }
      : {}),
    ...(parseColor(colorElement, theme, palette)
      ? { color: parseColor(colorElement, theme, palette) }
      : {}),
  };
  if (rich) {
    return {
      ...(style.bold ? { bold: true } : {}),
      ...(style.italic ? { italic: true } : {}),
      ...(style.strike ? { strike: true } : {}),
      ...(style.name ? { name: style.name } : {}),
      ...(style.size !== undefined ? { size: style.size } : {}),
      ...(style.underline ? { underline: style.underline } : {}),
      ...(style.vertAlign ? { vertAlign: style.vertAlign } : {}),
      ...(style.color ? { color: style.color } : {}),
    } as FontStyle;
  }
  return style;
};

const borderLines: Readonly<Record<string, BorderLine>> = {
  thin: "thin",
  medium: "medium",
  thick: "thick",
  dashed: "dashed",
  dotted: "dotted",
  double: "double",
  hair: "hair",
  mediumDashed: "mediumDashed",
  dashDot: "dashDot",
  mediumDashDot: "mediumDashDot",
  dashDotDot: "dashDotDot",
  mediumDashDotDot: "mediumDashDotDot",
  slantDashDot: "slantDashDot",
};

const readBorder = (
  element: Element | undefined,
  theme: readonly string[],
  palette: readonly string[],
): BorderEdge | undefined => {
  if (!element) return undefined;
  const line = borderLines[attr(element, "style") ?? ""];
  if (!line) return undefined;
  const color = parseColor(first(element, "color"), theme, palette);
  return color ? { line, color } : { line };
};

const readBorders = (
  document: Document | undefined,
  theme: readonly string[],
  palette: readonly string[],
): readonly BorderSet[] => {
  const borderElements = elements(first(document?.documentElement, "borders"), "border");
  return borderElements.map((border) => {
    const top = readBorder(first(border, "top"), theme, palette);
    const right = readBorder(first(border, "right"), theme, palette);
    const bottom = readBorder(first(border, "bottom"), theme, palette);
    const left = readBorder(first(border, "left"), theme, palette);
    const edges: BorderSet = {
      ...(top ? { top } : {}),
      ...(right ? { right } : {}),
      ...(bottom ? { bottom } : {}),
      ...(left ? { left } : {}),
    };
    return edges;
  });
};

export const defaultCellStyle: CellStyle = {
  numberFormat: "General",
  font: { bold: false, italic: false, strike: false },
  border: {},
  alignment: { wrap: false, indent: 0 },
};

export interface StyleTable {
  readonly cellStyles: readonly CellStyle[];
  readonly fontForRun: (element: Element | undefined) => Partial<FontStyle> | undefined;
  readonly resolve: (index: number | undefined) => CellStyle;
}

export const parseStyles = (document?: Document, theme: readonly string[] = []): StyleTable => {
  const root = document?.documentElement;
  const palette = indexedColors(document);
  const customFormats = new Map<number, string>();
  for (const format of elements(first(root, "numFmts"), "numFmt")) {
    const id = Number(attr(format, "numFmtId"));
    const code = attr(format, "formatCode");
    if (Number.isInteger(id) && code !== undefined) customFormats.set(id, code);
  }
  const numberFormats = new Map<number, string>(
    Object.entries(builtInFormats).map(([id, code]) => [Number(id), code]),
  );
  for (const [id, code] of customFormats) numberFormats.set(id, code);
  const fonts = elements(first(root, "fonts"), "font").map((font) =>
    readFont(font, theme, palette),
  );
  const fills = elements(first(root, "fills"), "fill").map((fill) => {
    const pattern = first(fill, "patternFill");
    if (attr(pattern, "patternType") !== "solid") return undefined;
    return parseColor(first(pattern, "fgColor"), theme, palette);
  });
  const borders = readBorders(document, theme, palette);
  const cellXfs = elements(first(root, "cellXfs"), "xf");
  const cellStyles = cellXfs.map((xf) => {
    const font = fonts[Number(attr(xf, "fontId") ?? 0)] ?? defaultCellStyle.font;
    const fill = fills[Number(attr(xf, "fillId") ?? 0)];
    const border = borders[Number(attr(xf, "borderId") ?? 0)] ?? {};
    const numFmtId = Number(attr(xf, "numFmtId") ?? 0);
    const alignment = first(xf, "alignment");
    const horizontal = attr(alignment, "horizontal");
    const vertical = attr(alignment, "vertical");
    const validHorizontal = [
      "left",
      "center",
      "right",
      "fill",
      "justify",
      "centerContinuous",
      "distributed",
    ].includes(horizontal ?? "");
    const validVertical = ["top", "center", "bottom", "justify", "distributed"].includes(
      vertical ?? "",
    );
    const wrap = attr(alignment, "wrapText");
    const indent = numberAttribute(alignment, "indent");
    return {
      numberFormat: numberFormats.get(numFmtId) ?? "General",
      font,
      ...(fill ? { fill } : {}),
      border,
      alignment: {
        ...(validHorizontal
          ? { horizontal: horizontal as CellStyle["alignment"]["horizontal"] }
          : {}),
        ...(validVertical ? { vertical: vertical as CellStyle["alignment"]["vertical"] } : {}),
        wrap: wrap === "1" || wrap === "true",
        indent: indent !== undefined && indent >= 0 ? indent : 0,
      },
    };
  });
  const resolve = (index: number | undefined): CellStyle => {
    return index === undefined
      ? (cellStyles[0] ?? defaultCellStyle)
      : (cellStyles[index] ?? defaultCellStyle);
  };
  return {
    cellStyles,
    fontForRun: (element) => {
      if (!element) return undefined;
      const style = readFont(element, theme, palette, true);
      return Object.keys(style).length ? style : undefined;
    },
    resolve,
  };
};
