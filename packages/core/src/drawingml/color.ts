import { hexToRgb, hslToRgb, rgbToHex, rgbToHsl, type Rgb } from "../common/color";
import { attr, children } from "../common/xml";

/** Resolves a scheme color name (`accent1`, `tx1`, `phClr`, ...) to `#rrggbb`. */
export type SchemeResolver = (name: string) => string | undefined;

const presetColors: Readonly<Record<string, string>> = {
  black: "#000000",
  white: "#ffffff",
  red: "#ff0000",
  green: "#008000",
  blue: "#0000ff",
  yellow: "#ffff00",
  gray: "#808080",
  grey: "#808080",
  orange: "#ffa500",
};

const colorElements = new Set(["srgbClr", "schemeClr", "sysClr", "prstClr", "scrgbClr", "hslClr"]);

const percent = (element: Element, name: string): number => {
  const value = Number(attr(element, name));
  return Number.isFinite(value) ? value / 100000 : 0;
};

const baseColor = (element: Element, scheme: SchemeResolver): string | undefined => {
  const val = attr(element, "val");
  switch (element.localName) {
    case "srgbClr":
      return val && /^[\da-f]{6}$/i.test(val) ? `#${val.toLowerCase()}` : undefined;
    case "sysClr": {
      const last = attr(element, "lastClr");
      return last && /^[\da-f]{6}$/i.test(last)
        ? `#${last.toLowerCase()}`
        : val === "window"
          ? "#ffffff"
          : "#000000";
    }
    case "schemeClr":
      return val ? scheme(val) : undefined;
    case "prstClr":
      return val ? presetColors[val] : undefined;
    case "scrgbClr":
      return rgbToHex([percent(element, "r"), percent(element, "g"), percent(element, "b")]);
    case "hslClr":
      return rgbToHex(
        hslToRgb([
          (Number(attr(element, "hue")) || 0) / 21600000,
          percent(element, "sat"),
          percent(element, "lum"),
        ]),
      );
  }
  return undefined;
};

const transform = (rgb: Rgb, alpha: number, modifier: Element): [Rgb, number] => {
  const value = percent(modifier, "val");
  switch (modifier.localName) {
    case "alpha":
      return [rgb, value];
    case "tint":
      return [rgb.map((channel) => channel * value + 1 - value) as unknown as Rgb, alpha];
    case "shade":
      return [rgb.map((channel) => channel * value) as unknown as Rgb, alpha];
    case "lumMod":
    case "lumOff":
    case "satMod": {
      const [hue, saturation, lightness] = rgbToHsl(rgb);
      return [
        hslToRgb(
          modifier.localName === "satMod"
            ? [hue, Math.min(1, saturation * value), lightness]
            : [
                hue,
                saturation,
                Math.min(
                  1,
                  modifier.localName === "lumMod" ? lightness * value : lightness + value,
                ),
              ],
        ),
        alpha,
      ];
    }
  }
  return [rgb, alpha];
};

/** Reads the color choice inside `parent` (e.g. `a:solidFill`) as `#rrggbb` or `#rrggbbaa`. */
export const parseColor = (
  parent: Element | undefined,
  scheme: SchemeResolver,
): string | undefined => {
  const element = parent && children(parent).find((child) => colorElements.has(child.localName));
  const base = element && baseColor(element, scheme);
  if (!base) return undefined;
  let rgb = hexToRgb(base);
  let alpha = 1;
  for (const modifier of children(element)) [rgb, alpha] = transform(rgb, alpha, modifier);
  const hex = rgbToHex(rgb);
  return alpha < 1
    ? `${hex}${Math.round(Math.max(0, alpha) * 255)
        .toString(16)
        .padStart(2, "0")}`
    : hex;
};
