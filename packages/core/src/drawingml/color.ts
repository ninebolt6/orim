import { hexToRgb, hslToRgb, rgbToHex, rgbToHsl, type Rgb } from "../common/color";
import { attr, children } from "../common/xml";

/** Resolves a scheme color name (`accent1`, `tx1`, `phClr`, ...) to `#rrggbb`. */
export type SchemeResolver = (name: string) => string | undefined;

/** `a:prstClr` names (ST_PresetColorVal) are the CSS named colors with `dk`/`lt`/`med` abbreviations. */
let presetColors: ReadonlyMap<string, string> | undefined;
const presetTable = (): ReadonlyMap<string, string> =>
  (presetColors ??= new Map(
    "aliceblue f0f8ff antiquewhite faebd7 aqua 00ffff aquamarine 7fffd4 azure f0ffff beige f5f5dc bisque ffe4c4 black 000000 blanchedalmond ffebcd blue 0000ff blueviolet 8a2be2 brown a52a2a burlywood deb887 cadetblue 5f9ea0 chartreuse 7fff00 chocolate d2691e coral ff7f50 cornflowerblue 6495ed cornsilk fff8dc crimson dc143c cyan 00ffff darkblue 00008b darkcyan 008b8b darkgoldenrod b8860b darkgray a9a9a9 darkgreen 006400 darkkhaki bdb76b darkmagenta 8b008b darkolivegreen 556b2f darkorange ff8c00 darkorchid 9932cc darkred 8b0000 darksalmon e9967a darkseagreen 8fbc8f darkslateblue 483d8b darkslategray 2f4f4f darkturquoise 00ced1 darkviolet 9400d3 deeppink ff1493 deepskyblue 00bfff dimgray 696969 dodgerblue 1e90ff firebrick b22222 floralwhite fffaf0 forestgreen 228b22 fuchsia ff00ff gainsboro dcdcdc ghostwhite f8f8ff gold ffd700 goldenrod daa520 gray 808080 green 008000 greenyellow adff2f honeydew f0fff0 hotpink ff69b4 indianred cd5c5c indigo 4b0082 ivory fffff0 khaki f0e68c lavender e6e6fa lavenderblush fff0f5 lawngreen 7cfc00 lemonchiffon fffacd lightblue add8e6 lightcoral f08080 lightcyan e0ffff lightgoldenrodyellow fafad2 lightgray d3d3d3 lightgreen 90ee90 lightpink ffb6c1 lightsalmon ffa07a lightseagreen 20b2aa lightskyblue 87cefa lightslategray 778899 lightsteelblue b0c4de lightyellow ffffe0 lime 00ff00 limegreen 32cd32 linen faf0e6 magenta ff00ff maroon 800000 mediumaquamarine 66cdaa mediumblue 0000cd mediumorchid ba55d3 mediumpurple 9370db mediumseagreen 3cb371 mediumslateblue 7b68ee mediumspringgreen 00fa9a mediumturquoise 48d1cc mediumvioletred c71585 midnightblue 191970 mintcream f5fffa mistyrose ffe4e1 moccasin ffe4b5 navajowhite ffdead navy 000080 oldlace fdf5e6 olive 808000 olivedrab 6b8e23 orange ffa500 orangered ff4500 orchid da70d6 palegoldenrod eee8aa palegreen 98fb98 paleturquoise afeeee palevioletred db7093 papayawhip ffefd5 peachpuff ffdab9 peru cd853f pink ffc0cb plum dda0dd powderblue b0e0e6 purple 800080 red ff0000 rosybrown bc8f8f royalblue 4169e1 saddlebrown 8b4513 salmon fa8072 sandybrown f4a460 seagreen 2e8b57 seashell fff5ee sienna a0522d silver c0c0c0 skyblue 87ceeb slateblue 6a5acd slategray 708090 snow fffafa springgreen 00ff7f steelblue 4682b4 tan d2b48c teal 008080 thistle d8bfd8 tomato ff6347 turquoise 40e0d0 violet ee82ee wheat f5deb3 white ffffff whitesmoke f5f5f5 yellow ffff00 yellowgreen 9acd32"
      .split(" ")
      .flatMap((item, index, list) => (index % 2 ? [] : [[item, `#${list[index + 1]}`] as const])),
  ));

const presetColor = (name: string): string | undefined =>
  presetTable().get(
    name
      .toLowerCase()
      .replace(/^dk/, "dark")
      .replace(/^lt/, "light")
      .replace(/^med(?!ium)/, "medium")
      .replace("grey", "gray"),
  );

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
      return val ? presetColor(val) : undefined;
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

const toLinear = (channel: number): number =>
  channel <= 0.04045 ? channel / 12.92 : ((channel + 0.055) / 1.055) ** 2.4;
const fromLinear = (channel: number): number =>
  channel <= 0.0031308 ? channel * 12.92 : 1.055 * channel ** (1 / 2.4) - 0.055;
/** Tint and shade operate on linear RGB, as PowerPoint renders them. */
const linear = (rgb: Rgb, map: (channel: number) => number): Rgb =>
  rgb.map((channel) => fromLinear(map(toLinear(channel)))) as unknown as Rgb;

const transform = (rgb: Rgb, alpha: number, modifier: Element): [Rgb, number] => {
  const value = percent(modifier, "val");
  switch (modifier.localName) {
    case "alpha":
      return [rgb, value];
    case "tint":
      return [linear(rgb, (channel) => channel * value + 1 - value), alpha];
    case "shade":
      return [linear(rgb, (channel) => channel * value), alpha];
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
  const hex = rgbToHex(rgb.map((channel) => Math.min(1, Math.max(0, channel))) as unknown as Rgb);
  return alpha < 1
    ? `${hex}${Math.round(Math.max(0, alpha) * 255)
        .toString(16)
        .padStart(2, "0")}`
    : hex;
};
