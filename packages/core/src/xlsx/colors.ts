import { attr, children, descendants, first } from "../common/xml";

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

const normalizeHex = (value: string | undefined): string | undefined => {
  if (!value || !/^(?:[0-9a-f]{6}|[0-9a-f]{8})$/i.test(value)) return undefined;
  return `#${value.slice(-6).toLowerCase()}`;
};

const tinted = (color: string, tint: number): string => {
  const red = Number.parseInt(color.slice(1, 3), 16) / 255;
  const green = Number.parseInt(color.slice(3, 5), 16) / 255;
  const blue = Number.parseInt(color.slice(5, 7), 16) / 255;
  const maximum = Math.max(red, green, blue);
  const minimum = Math.min(red, green, blue);
  let hue = 0;
  let saturation = 0;
  let lightness = (maximum + minimum) / 2;
  if (maximum !== minimum) {
    const difference = maximum - minimum;
    saturation =
      lightness > 0.5 ? difference / (2 - maximum - minimum) : difference / (maximum + minimum);
    if (maximum === red) hue = (green - blue) / difference + (green < blue ? 6 : 0);
    else if (maximum === green) hue = (blue - red) / difference + 2;
    else hue = (red - green) / difference + 4;
    hue /= 6;
  }
  lightness = tint < 0 ? lightness * (1 + tint) : lightness * (1 - tint) + tint;
  let tintedRed: number;
  let tintedGreen: number;
  let tintedBlue: number;
  if (saturation === 0) {
    tintedRed = lightness;
    tintedGreen = lightness;
    tintedBlue = lightness;
  } else {
    const upper =
      lightness < 0.5
        ? lightness * (1 + saturation)
        : lightness + saturation - lightness * saturation;
    const lower = 2 * lightness - upper;
    const channel = (offset: number): number => {
      let position = hue + offset;
      if (position < 0) position++;
      if (position > 1) position--;
      if (position < 1 / 6) return lower + (upper - lower) * 6 * position;
      if (position < 1 / 2) return upper;
      if (position < 2 / 3) return lower + (upper - lower) * (2 / 3 - position) * 6;
      return lower;
    };
    tintedRed = channel(1 / 3);
    tintedGreen = channel(0);
    tintedBlue = channel(-1 / 3);
  }
  return `#${[tintedRed, tintedGreen, tintedBlue]
    .map((channel) =>
      Math.round(channel * 255)
        .toString(16)
        .padStart(2, "0"),
    )
    .join("")}`;
};

export const themeColors = (document: Document | undefined): readonly string[] => {
  if (!document) return [];
  const scheme = descendants(document, "clrScheme")[0];
  if (!scheme) return [];
  const names = [
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
  return names.map((name) => {
    const slot = children(scheme).find((element) => element.localName === name);
    const color = slot && children(slot)[0];
    return (
      normalizeHex(
        color?.localName === "sysClr"
          ? attr(color, "lastClr")
          : color?.localName === "srgbClr"
            ? attr(color, "val")
            : undefined,
      ) ?? ""
    );
  });
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
