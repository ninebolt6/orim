import { attr, children, elements, first } from "../common/xml";
import { parseColor, type SchemeResolver } from "./color";
import { angle, emu, fraction } from "./units";

export interface GradientStop {
  /** Offset along the gradient, 0–1. */
  readonly position: number;
  readonly color: string;
}

export interface Crop {
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
}

export type Fill =
  | { readonly type: "none" }
  | { readonly type: "solid"; readonly color: string }
  | {
      readonly type: "gradient";
      readonly radial: boolean;
      /** Direction in degrees, clockwise from left-to-right. */
      readonly angle: number;
      readonly stops: readonly GradientStop[];
    }
  | { readonly type: "image"; readonly part: string; readonly crop?: Crop };

export interface LineStyle {
  /** Width in points. */
  readonly width?: number;
  readonly fill?: Fill;
  readonly dash?: string;
  readonly head?: string;
  readonly tail?: string;
}

export interface FillOptions {
  readonly scheme: SchemeResolver;
  /** Resolves an `a:blipFill` to its media part. */
  readonly image?: (blipFill: Element) => string | undefined;
  /** Fill used for `a:grpFill`. */
  readonly group?: Fill;
}

const fillNames = new Set(["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"]);

/** The fill choice element directly inside `parent` (e.g. `p:spPr`). */
export const fillElement = (parent: Element | undefined): Element | undefined =>
  parent && children(parent).find((child) => fillNames.has(child.localName));

export const parseCrop = (blipFill: Element): Crop | undefined => {
  const rect = first(blipFill, "srcRect");
  if (!rect) return undefined;
  const side = (name: string) => fraction(attr(rect, name)) ?? 0;
  const crop = { left: side("l"), top: side("t"), right: side("r"), bottom: side("b") };
  return Object.values(crop).some(Boolean) ? crop : undefined;
};

/** Parses a fill choice element. `undefined` means "not specified" (inherit). */
export const parseFill = (element: Element | undefined, options: FillOptions): Fill | undefined => {
  switch (element?.localName) {
    case "noFill":
      return { type: "none" };
    case "solidFill": {
      const color = parseColor(element, options.scheme);
      return color ? { type: "solid", color } : undefined;
    }
    case "pattFill": {
      const color = parseColor(first(element, "fgClr"), options.scheme);
      return color ? { type: "solid", color } : undefined;
    }
    case "gradFill": {
      const stops = elements(first(element, "gsLst"), "gs")
        .flatMap((stop) => {
          const color = parseColor(stop, options.scheme);
          return color ? [{ position: fraction(attr(stop, "pos")) ?? 0, color }] : [];
        })
        .sort((left, right) => left.position - right.position);
      if (!stops.length) return undefined;
      return {
        type: "gradient",
        radial: !!first(element, "path"),
        angle: angle(attr(first(element, "lin"), "ang")) ?? 0,
        stops,
      };
    }
    case "blipFill": {
      const part = options.image?.(element);
      return part ? { type: "image", part, crop: parseCrop(element) } : undefined;
    }
    case "grpFill":
      return options.group;
  }
  return undefined;
};

/** Parses `a:ln`; absent attributes stay `undefined` so callers can layer inheritance. */
export const parseLine = (
  element: Element | undefined,
  options: FillOptions,
): LineStyle | undefined => {
  if (!element) return undefined;
  const end = (name: string) => {
    const type = attr(first(element, name), "type");
    return type && type !== "none" ? type : undefined;
  };
  return {
    width: emu(attr(element, "w")),
    fill: parseFill(fillElement(element), options),
    dash: attr(first(element, "prstDash"), "val"),
    head: end("headEnd"),
    tail: end("tailEnd"),
  };
};

export const mergeLine = (
  base: LineStyle | undefined,
  override: LineStyle | undefined,
): LineStyle | undefined => {
  if (!base || !override) return override ?? base;
  return {
    width: override.width ?? base.width,
    fill: override.fill ?? base.fill,
    dash: override.dash ?? base.dash,
    head: override.head ?? base.head,
    tail: override.tail ?? base.tail,
  };
};
