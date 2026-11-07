import { hexToRgb, hslToRgb, normalizeHex, rgbToHex, rgbToHsl } from "../common/color";
import { attr, children, first } from "../common/xml";
import { themeColorScheme } from "../drawingml/theme";

export const defaultIndexedPalette = [
  "#000000",
  "#ffffff",
  "#ff0000",
  "#00ff00",
  "#0000ff",
  "#ffff00",
  "#ff00ff",
  "#00ffff",
  "#000000",
  "#ffffff",
  "#ff0000",
  "#00ff00",
  "#0000ff",
  "#ffff00",
  "#ff00ff",
  "#00ffff",
  "#800000",
  "#008000",
  "#000080",
  "#808000",
  "#800080",
  "#008080",
  "#c0c0c0",
  "#808080",
  "#9999ff",
  "#993366",
  "#ffffcc",
  "#ccffff",
  "#660066",
  "#ff8080",
  "#0066cc",
  "#ccccff",
  "#000080",
  "#ff00ff",
  "#ffff00",
  "#00ffff",
  "#800080",
  "#800000",
  "#008080",
  "#0000ff",
  "#00ccff",
  "#ccffff",
  "#ccffcc",
  "#ffff99",
  "#99ccff",
  "#ff99cc",
  "#cc99ff",
  "#ffcc99",
  "#3366ff",
  "#33cccc",
  "#99cc00",
  "#ffcc00",
  "#ff9900",
  "#ff6600",
  "#666699",
  "#969696",
  "#003366",
  "#339966",
  "#003300",
  "#333300",
  "#993300",
  "#993366",
  "#333399",
  "#333333",
];

const tinted = (color: string, tint: number): string => {
  const [hue, saturation, lightness] = rgbToHsl(hexToRgb(color));
  return rgbToHex(
    hslToRgb([hue, saturation, tint < 0 ? lightness * (1 + tint) : lightness * (1 - tint) + tint]),
  );
};

const themeSlots = [
  "lt1",
  "dk1",
  "lt2",
  "dk2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink",
];

/** Theme colors in SpreadsheetML index order (light/dark slots swapped). */
export const themeColors = (document: Document | undefined): readonly string[] => {
  const scheme = themeColorScheme(document);
  return Object.keys(scheme).length ? themeSlots.map((name) => scheme[name] ?? "") : [];
};

export const indexedColors = (document?: Document): readonly string[] => {
  const palette = document && first(first(document.documentElement, "colors"), "indexedColors");
  if (!palette) return defaultIndexedPalette;
  const custom = children(palette).map((element) => normalizeHex(attr(element, "rgb")) ?? "");
  return custom.length ? custom : defaultIndexedPalette;
};

export const parseColor = (
  element: Element | undefined,
  theme: readonly string[] = [],
  indexed: readonly string[] = defaultIndexedPalette,
): string | undefined => {
  if (!element || attr(element, "auto") === "1") return undefined;
  const rgb = normalizeHex(attr(element, "rgb"));
  const themeIndex = Number(attr(element, "theme"));
  const indexedIndex = Number(attr(element, "indexed"));
  const base =
    rgb ??
    (Number.isInteger(themeIndex) && themeIndex >= 0 ? theme[themeIndex] : undefined) ??
    (Number.isInteger(indexedIndex) && indexedIndex >= 0
      ? (indexed[indexedIndex] ?? (indexedIndex === 64 ? "#000000" : undefined))
      : undefined) ??
    (indexedIndex === 65 ? "#ffffff" : undefined);
  if (!base || !/^#[0-9a-f]{6}$/i.test(base)) return undefined;
  const tint = Number(attr(element, "tint") ?? 0);
  return Number.isFinite(tint) && tint !== 0 ? tinted(base, Math.max(-1, Math.min(1, tint))) : base;
};
