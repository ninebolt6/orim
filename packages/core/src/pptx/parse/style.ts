import { attr } from "../../common/xml";
import { blipPart } from "../../drawingml/blip";
import { parseColor } from "../../drawingml/color";
import {
  parseFill,
  parseLine,
  type Fill,
  type FillOptions,
  type LineStyle,
} from "../../drawingml/fill";
import type { RunStyle } from "../model";
import { defined, withPlaceholderColor, type ShapeContext, type SlideContext } from "./context";

export const fillOptions = (context: ShapeContext, group?: Fill): FillOptions => ({
  scheme: context.scheme,
  image: (blipFill) => blipPart(blipFill, context.part.rels, context.part.path, context.part.opc),
  group,
});

const referenced = (reference: Element, context: SlideContext) => {
  const index = Number(attr(reference, "idx")) || 0;
  return {
    index,
    scheme: withPlaceholderColor(context, parseColor(reference, context.scheme)).scheme,
  };
};

/** Resolves `fillRef`/`bgRef` against the theme's `fmtScheme`. */
export const referenceFill = (
  reference: Element | undefined,
  context: SlideContext,
): Fill | undefined => {
  if (!reference) return undefined;
  const { index, scheme } = referenced(reference, context);
  if (!index) return { type: "none" };
  const source =
    index >= 1001 ? context.formats.backgrounds[index - 1001] : context.formats.fills[index - 1];
  return parseFill(source, { scheme });
};

export const referenceLine = (
  reference: Element | undefined,
  context: SlideContext,
): LineStyle | undefined => {
  if (!reference) return undefined;
  const { index, scheme } = referenced(reference, context);
  if (!index) return { fill: { type: "none" } };
  return parseLine(context.formats.lines[index - 1], { scheme });
};

export const referenceFont = (
  reference: Element | undefined,
  context: SlideContext,
): Partial<RunStyle> | undefined => {
  if (!reference) return undefined;
  const index = attr(reference, "idx");
  return defined({
    color: parseColor(reference, context.scheme),
    font:
      index === "major" ? context.fonts.major : index === "minor" ? context.fonts.minor : undefined,
  });
};
