export type Rgb = readonly [number, number, number];
export type Hsl = readonly [number, number, number];

/** Normalizes `RRGGBB` or `AARRGGBB` to `#rrggbb`. */
export const normalizeHex = (value: string | undefined): string | undefined =>
  value && /^(?:[\da-f]{6}|[\da-f]{8})$/i.test(value)
    ? `#${value.slice(-6).toLowerCase()}`
    : undefined;

export const hexToRgb = (hex: string): Rgb => [
  Number.parseInt(hex.slice(1, 3), 16) / 255,
  Number.parseInt(hex.slice(3, 5), 16) / 255,
  Number.parseInt(hex.slice(5, 7), 16) / 255,
];

/** Channels must be within 0–1. */
export const rgbToHex = (rgb: Rgb): string =>
  `#${rgb
    .map((channel) =>
      Math.round(channel * 255)
        .toString(16)
        .padStart(2, "0"),
    )
    .join("")}`;

export const rgbToHsl = ([red, green, blue]: Rgb): Hsl => {
  const maximum = Math.max(red, green, blue);
  const minimum = Math.min(red, green, blue);
  const lightness = (maximum + minimum) / 2;
  if (maximum === minimum) return [0, 0, lightness];
  const difference = maximum - minimum;
  const saturation =
    lightness > 0.5 ? difference / (2 - maximum - minimum) : difference / (maximum + minimum);
  const hue =
    maximum === red
      ? (green - blue) / difference + (green < blue ? 6 : 0)
      : maximum === green
        ? (blue - red) / difference + 2
        : (red - green) / difference + 4;
  return [hue / 6, saturation, lightness];
};

export const hslToRgb = ([hue, saturation, lightness]: Hsl): Rgb => {
  if (saturation === 0) return [lightness, lightness, lightness];
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
  return [channel(1 / 3), channel(0), channel(-1 / 3)];
};
