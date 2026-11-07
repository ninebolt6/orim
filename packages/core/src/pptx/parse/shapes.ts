import { attr, children, first } from "../../common/xml";
import { blipPart } from "../../drawingml/blip";
import {
  fillElement,
  mergeLine,
  parseCrop,
  parseFill,
  parseLine,
  type Fill,
} from "../../drawingml/fill";
import { emu } from "../../drawingml/units";
import type { Bounds, SlideElement, Transform } from "../model";
import { defined, type PartContext, type ShapeContext } from "./context";
import { parseGeometry, parseTransform } from "./geometry";
import { fillOptions, referenceFill, referenceFont, referenceLine } from "./style";
import { parseTable } from "./table";
import {
  linkTarget,
  parseBodyProps,
  parseListStyle,
  parseTextBody,
  uniformListStyle,
  type ListStyle,
} from "./text";

export interface Placeholder {
  readonly type: string;
  readonly index?: string;
}

/** A layout or master placeholder that a slide shape inherits from. */
export interface InheritedShape {
  readonly element: Element;
  readonly part: PartContext;
}

export interface Inheritance {
  /** Inherited placeholders ordered from master to layout. */
  readonly sources: readonly InheritedShape[];
  readonly textStyle?: ListStyle;
}

export interface TreeOptions {
  /** Layouts and masters draw only their non-placeholder shapes. */
  readonly template: boolean;
  readonly inherit: (placeholder: Placeholder) => Inheritance;
  readonly defaultText: ListStyle;
}

const nonVisual = (element: Element): Element | undefined =>
  children(element).find((child) => child.localName.startsWith("nv"));

export const placeholderOf = (element: Element): Placeholder | undefined => {
  const placeholder = first(first(nonVisual(element), "nvPr"), "ph");
  if (!placeholder) return undefined;
  const index = attr(placeholder, "idx");
  const type = attr(placeholder, "type") ?? "body";
  return index === undefined ? { type } : { type, index };
};

interface Layer {
  readonly element: Element;
  readonly context: ShapeContext;
}

const properties = (element: Element): Element | undefined =>
  children(element).find((child) => child.localName === "spPr" || child.localName === "grpSpPr");

const highest = <T>(
  layers: readonly Layer[],
  read: (layer: Layer) => T | undefined,
): T | undefined =>
  layers.reduceRight<T | undefined>((result, item) => result ?? read(item), undefined);

const groupBounds = (xfrm: Element | undefined, transform: Transform): Bounds => {
  const offset = first(xfrm, "chOff");
  const extent = first(xfrm, "chExt");
  return offset && extent
    ? {
        x: emu(attr(offset, "x")) ?? 0,
        y: emu(attr(offset, "y")) ?? 0,
        width: emu(attr(extent, "cx")) ?? 0,
        height: emu(attr(extent, "cy")) ?? 0,
      }
    : transform;
};

const parseShape = (
  element: Element,
  context: ShapeContext,
  options: TreeOptions,
  group: Fill | undefined,
): SlideElement[] => {
  const placeholder = placeholderOf(element);
  if (placeholder && options.template) return [];
  const inheritance = placeholder ? options.inherit(placeholder) : { sources: [] };
  const layers: readonly Layer[] = [
    ...inheritance.sources.map((source) => ({
      element: source.element,
      context: { ...context, part: source.part },
    })),
    { element, context },
  ];
  const transform = highest(layers, (item) =>
    parseTransform(first(properties(item.element), "xfrm")),
  );
  if (!transform) return [];
  const info = first(nonVisual(element), "cNvPr");
  const base = defined({
    id: attr(info, "id") ?? "",
    name: attr(info, "name") ?? "",
    transform,
    href: linkTarget(info, context),
  }) as { id: string; name: string; transform: Transform; href?: string };
  const geometry = highest(layers, (item) => parseGeometry(properties(item.element))) ?? {
    preset: "rect",
    adjust: {},
  };
  const style = highest(layers, (item) => first(item.element, "style"));
  const line = layers.reduce(
    (result, item) => mergeLine(result, parseLineOf(properties(item.element), item.context, group)),
    referenceLine(first(style, "lnRef"), context),
  );
  if (element.localName === "pic") {
    const blipFill = first(element, "blipFill");
    const part =
      blipFill && blipPart(blipFill, context.part.rels, context.part.path, context.part.opc);
    if (!part) return [];
    return [
      defined({
        kind: "picture" as const,
        ...base,
        geometry,
        part,
        crop: parseCrop(blipFill),
        line,
        description: attr(info, "descr") || undefined,
      }) as SlideElement,
    ];
  }
  const fill =
    highest(layers, (item) => {
      const choice = fillElement(properties(item.element));
      return choice && parseFill(choice, fillOptions(item.context, group));
    }) ?? referenceFill(first(style, "fillRef"), context);
  const font = referenceFont(first(style, "fontRef"), context);
  const text = parseTextBody(
    first(element, "txBody"),
    [
      placeholder ? inheritance.textStyle : options.defaultText,
      font && uniformListStyle(font),
      ...layers
        .slice(0, -1)
        .map((item) =>
          parseListStyle(first(first(item.element, "txBody"), "lstStyle"), item.context),
        ),
    ],
    layers
      .slice(0, -1)
      .map((item) => parseBodyProps(first(first(item.element, "txBody"), "bodyPr"))),
    context,
  );
  return [defined({ kind: "shape" as const, ...base, geometry, fill, line, text }) as SlideElement];
};

const parseLineOf = (
  shapeProperties: Element | undefined,
  context: ShapeContext,
  group: Fill | undefined,
) => {
  const line = first(shapeProperties, "ln");
  return line ? parseLine(line, fillOptions(context, group)) : undefined;
};

export const parseTree = (
  parent: Element,
  context: ShapeContext,
  options: TreeOptions,
  group?: Fill,
): SlideElement[] =>
  children(parent).flatMap((child): SlideElement[] => {
    if (attr(first(nonVisual(child), "cNvPr"), "hidden") === "1") return [];
    switch (child.localName) {
      case "sp":
      case "cxnSp":
      case "pic":
        return parseShape(child, context, options, group);
      case "grpSp": {
        const groupProperties = first(child, "grpSpPr");
        const xfrm = first(groupProperties, "xfrm");
        const transform = parseTransform(xfrm);
        if (!transform) return [];
        const info = first(nonVisual(child), "cNvPr");
        const fill = parseFill(fillElement(groupProperties), fillOptions(context, group)) ?? group;
        return [
          {
            kind: "group",
            id: attr(info, "id") ?? "",
            name: attr(info, "name") ?? "",
            transform,
            childBounds: groupBounds(xfrm, transform),
            children: parseTree(child, context, options, fill),
          },
        ];
      }
      case "graphicFrame": {
        if (placeholderOf(child) && options.template) return [];
        const transform = parseTransform(first(child, "xfrm"));
        const info = first(nonVisual(child), "cNvPr");
        const table =
          transform &&
          parseTable(
            child,
            { id: attr(info, "id") ?? "", name: attr(info, "name") ?? "", transform },
            context,
          );
        return table ? [table] : [];
      }
      case "AlternateContent": {
        const choice = first(child, "Fallback") ?? first(child, "Choice");
        return choice ? parseTree(choice, context, options, group) : [];
      }
    }
    return [];
  });
