import type { ChangeMark, TextStyle } from "../model";
import { validColor } from "../../common/sanitize";

const palette = ["#137333", "#c5221f", "#185abc", "#a142f4", "#b06000"];
const highlightColors: Record<string, string> = {
  black: "#000",
  blue: "#00f",
  cyan: "#0ff",
  darkBlue: "#000080",
  darkCyan: "#008080",
  darkGray: "#808080",
  darkGreen: "#008000",
  darkMagenta: "#800080",
  darkRed: "#800000",
  darkYellow: "#808000",
  green: "#0f0",
  lightGray: "#c0c0c0",
  magenta: "#f0f",
  red: "#f00",
  white: "#fff",
  yellow: "#ff0",
};

export const changeColor = (change: ChangeMark, colors: Map<string, string>): string => {
  let value = colors.get(change.author);
  if (!value) {
    value = palette[colors.size % palette.length];
    colors.set(change.author, value);
  }
  return value;
};

export const fontStack = (name: string | undefined): string | undefined => {
  if (!name) return undefined;
  const names = [
    ...new Set(
      name
        .split(",")
        .map((item) => item.trim())
        .filter(
          (item) => Boolean(item) && !/["'\\;@<>]|\/\*|\b(?:url|image|expression)\s*\(/i.test(item),
        ),
    ),
  ];
  if (!names.length) return undefined;
  const families = names
    .map((item) =>
      ["serif", "sans-serif", "monospace", "cursive", "fantasy", "system-ui"].includes(
        item.toLowerCase(),
      )
        ? item
        : `"${item}"`,
    )
    .join(", ");
  const fallback = /明朝|Mincho/i.test(name)
    ? '"Yu Mincho", "YuMincho", "Hiragino Mincho ProN", "Noto Serif JP", "Noto Serif CJK JP", serif'
    : /ゴシック|Gothic|Meiryo|メイリオ/i.test(name)
      ? '"Yu Gothic", "YuGothic", "Hiragino Sans", "Noto Sans JP", "Noto Sans CJK JP", sans-serif'
      : /Times|Century|Cambria|Georgia|Mincho|Serif/i.test(name)
        ? "serif"
        : "sans-serif";
  return `${families}, ${fallback}`;
};

export const applyTextStyle = (
  element: HTMLElement,
  style: TextStyle | undefined,
  revisionMarked = false,
): void => {
  if (!style) return;
  if (style.bold !== undefined) element.style.fontWeight = style.bold ? "700" : "400";
  if (style.italic !== undefined) element.style.fontStyle = style.italic ? "italic" : "normal";
  if (style.underline === "none") {
    if (!revisionMarked) element.style.textDecorationLine = "none";
  } else if (style.underline) {
    element.style.textDecorationLine = "underline";
    element.style.textDecorationStyle =
      style.underline === "double" ||
      style.underline === "dotted" ||
      style.underline === "dash" ||
      style.underline === "wave"
        ? style.underline === "dash"
          ? "dashed"
          : style.underline
        : "solid";
  }
  if (style.strike) {
    const current = element.style.textDecorationLine;
    element.style.textDecorationLine = `${current === "none" ? "" : current} line-through`.trim();
  }
  if (style.doubleStrike) element.style.textDecoration = "line-through double";
  const color = validColor(style.color);
  if (color) element.style.color = color;
  if (style.highlight)
    element.style.backgroundColor =
      highlightColors[style.highlight] ?? validColor(style.highlight) ?? "";
  else if (validColor(style.backgroundColor))
    element.style.backgroundColor = style.backgroundColor!;
  if (style.fontSize && Number.isFinite(style.fontSize))
    element.style.fontSize = `${style.fontSize}pt`;
  const family = fontStack(style.fontFamily);
  if (family) element.style.fontFamily = family;
  if (style.verticalAlign) {
    element.style.verticalAlign = style.verticalAlign === "superscript" ? "super" : "sub";
    element.style.fontSize = style.fontSize ? `${style.fontSize * 0.65}pt` : "65%";
  }
  if (style.caps) element.style.textTransform = "uppercase";
  if (style.smallCaps) element.style.fontVariant = "small-caps";
  if (style.hidden) element.style.display = "none";
};
