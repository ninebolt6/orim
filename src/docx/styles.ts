import { attr, descendants, first } from "../opc";
import type { ParagraphProperties, TextStyle } from "./model";

type Props = Record<string, string | number | boolean | undefined>;
type BorderSides = Partial<
  Record<"top" | "right" | "bottom" | "left" | "insideH" | "insideV", string>
>;
type CellMargins = Partial<Record<"top" | "right" | "bottom" | "left", number>>;

export interface TableStyleProperties {
  borders: BorderSides;
  margins: CellMargins;
  cellShading?: string;
}

interface Style {
  basedOn?: string;
  paragraph: Props;
  run: Props;
  numbering: { numId?: string; level?: number };
}

export interface StyleSheet {
  paragraph(styleId?: string, direct?: Partial<ParagraphProperties>): ParagraphProperties;
  run(paragraphStyle?: string, runStyle?: string, direct?: Partial<TextStyle>): TextStyle;
  numbering(styleId?: string): { numId?: string; level?: number };
  table(styleId?: string): TableStyleProperties;
}

function value(element: Element | undefined): string | undefined {
  return attr(element, "val");
}

function bool(element: Element | undefined): boolean | undefined {
  const val = value(element);
  if (val === undefined) return element ? true : undefined;
  return !["0", "false", "off", "no"].includes(val.toLowerCase());
}

function points(val: string | undefined, divisor: number): number | undefined {
  if (val === undefined) return undefined;
  const number = Number(val);
  return Number.isFinite(number) ? number / divisor : undefined;
}

function twips(val: string | undefined): number | undefined {
  return points(val, 20);
}

function safeColor(val: string | undefined): string | undefined {
  return val && /^[\da-f]{6}$/i.test(val) ? `#${val}` : undefined;
}

export function border(element: Element | undefined): string | undefined {
  if (!element) return undefined;
  const value = attr(element, "val");
  if (value === "nil" || value === "none") return "none";
  const requestedSize = Number(attr(element, "sz") ?? 4) / 8;
  const size = Math.max(0.25, Number.isFinite(requestedSize) ? requestedSize : 0.5);
  const style =
    value === "double"
      ? "double"
      : value === "dotted"
        ? "dotted"
        : value === "dashed" || value === "dashSmallGap"
          ? "dashed"
          : "solid";
  const color = safeColor(attr(element, "color")) ?? "currentColor";
  return `${size}pt ${style} ${color}`;
}

function tableStyleProperties(node: Element): TableStyleProperties {
  const tableProperties = first(node, "tblPr");
  const borderProperties = first(tableProperties ?? node, "tblBorders");
  const borders: BorderSides = {};
  for (const side of ["top", "right", "bottom", "left", "insideH", "insideV"] as const) {
    const value = border(first(borderProperties ?? node, side));
    if (value !== undefined) borders[side] = value;
  }
  const marginProperties = first(tableProperties ?? node, "tblCellMar");
  const margins: CellMargins = {};
  for (const side of ["top", "right", "bottom", "left"] as const) {
    const value = first(marginProperties ?? node, side);
    const amount = points(attr(value, "w"), 20);
    if (amount !== undefined) margins[side] = amount;
  }
  const shading = first(first(node, "tcPr") ?? node, "shd");
  const fill = attr(shading, "fill");
  return { borders, margins, cellShading: fill && fill !== "auto" ? safeColor(fill) : undefined };
}

type Theme = { major?: string; minor?: string; majorEastAsia?: string; minorEastAsia?: string };

function runProperties(parent: Element | undefined, theme: Theme): Props {
  if (!parent) return {};
  const fonts = first(parent, "rFonts");
  const latinTheme = attr(fonts, "asciiTheme") ?? attr(fonts, "hAnsiTheme");
  const eastAsiaTheme = attr(fonts, "eastAsiaTheme");
  const family = [
    attr(fonts, "ascii") ??
      attr(fonts, "hAnsi") ??
      (latinTheme?.startsWith("major") ? theme.major : theme.minor),
    attr(fonts, "eastAsia") ??
      (eastAsiaTheme?.startsWith("major") ? theme.majorEastAsia : theme.minorEastAsia),
  ]
    .filter((item, index, all): item is string => !!item && all.indexOf(item) === index)
    .join(",");
  const size = points(value(first(parent, "sz")), 2);
  const color = value(first(parent, "color"));
  const highlight = value(first(parent, "highlight"));
  const vertAlign = value(first(parent, "vertAlign"));
  const underline = value(first(parent, "u"));
  const props: Props = {
    bold: bool(first(parent, "b")),
    italic: bool(first(parent, "i")),
    underline:
      underline === "none" ? undefined : (underline ?? (first(parent, "u") ? "single" : undefined)),
    strike: bool(first(parent, "strike")),
    doubleStrike: bool(first(parent, "dstrike")),
    color: color === "auto" ? undefined : safeColor(color),
    highlight: highlight && highlight !== "none" ? highlight : undefined,
    fontFamily: family || undefined,
    fontSize: size,
    verticalAlign: vertAlign === "superscript" || vertAlign === "subscript" ? vertAlign : undefined,
    caps: bool(first(parent, "caps")),
    smallCaps: bool(first(parent, "smallCaps")),
    hidden: bool(first(parent, "vanish")),
  };
  const shading = first(parent, "shd");
  const fill = attr(shading, "fill");
  if (fill && fill !== "auto") props.backgroundColor = safeColor(fill);
  return props;
}

export function paragraphProperties(parent: Element | undefined): Partial<ParagraphProperties> {
  if (!parent) return {};
  const spacing = first(parent, "spacing");
  const indent = first(parent, "ind");
  const jc = value(first(parent, "jc"));
  const shading = first(parent, "shd");
  const fill = attr(shading, "fill");
  const line = attr(spacing, "line");
  const lineRule = attr(spacing, "lineRule") ?? (line === undefined ? undefined : "auto");
  const props: Partial<ParagraphProperties> = {
    align: jc,
    before: twips(attr(spacing, "before")),
    after: twips(attr(spacing, "after")),
    line: lineRule === "auto" ? points(line, 240) : twips(line),
    lineRule,
    left: twips(attr(indent, "left") ?? attr(indent, "start")),
    right: twips(attr(indent, "right") ?? attr(indent, "end")),
    firstLine: twips(attr(indent, "firstLine")),
    hanging: twips(attr(indent, "hanging")),
    contextualSpacing: bool(first(parent, "contextualSpacing")),
    pageBreakBefore: bool(first(parent, "pageBreakBefore")),
    keepNext: bool(first(parent, "keepNext")),
    snapToGrid: bool(first(parent, "snapToGrid")),
    shading: fill && fill !== "auto" ? safeColor(fill) : undefined,
  };
  const borders = first(parent, "pBdr");
  if (borders) {
    const values: NonNullable<ParagraphProperties["borders"]> = {};
    const spacing: NonNullable<ParagraphProperties["borderSpacing"]> = {};
    for (const side of ["top", "right", "bottom", "left"] as const) {
      const value = first(borders, side);
      const parsed = border(value);
      if (parsed !== undefined) values[side] = parsed;
      const amount = points(attr(value, "space"), 1);
      if (amount !== undefined) spacing[side] = amount;
    }
    if (Object.keys(values).length) props.borders = values;
    if (Object.keys(spacing).length) props.borderSpacing = spacing;
  }
  return props;
}

function merge<T>(...layers: Array<Props | undefined>): T {
  const out: Props = {};
  for (const layer of layers) {
    if (!layer) continue;
    for (const [key, item] of Object.entries(layer)) {
      if (item === undefined) continue;
      if ((key === "borders" || key === "borderSpacing") && item && typeof item === "object")
        out[key] = { ...(out[key] as object | undefined), ...item };
      else out[key] = item;
    }
  }
  return out as T;
}

export function parseStyles(doc: Document | undefined, theme: Theme = {}): StyleSheet {
  const styles = new Map<string, Style>();
  const tableStyles = new Map<string, { basedOn?: string; properties: TableStyleProperties }>();
  let defaultTableStyle: string | undefined;
  let defaultParagraph: Props = {};
  let defaultRun: Props = {};
  if (doc) {
    const defaults = descendants(doc, "docDefaults")[0];
    defaultParagraph = paragraphProperties(
      first(first(defaults ?? doc, "pPrDefault") ?? doc, "pPr"),
    );
    defaultRun = runProperties(first(first(defaults ?? doc, "rPrDefault") ?? doc, "rPr"), theme);
    for (const node of descendants(doc, "style")) {
      const id = attr(node, "styleId") ?? node.getAttribute("styleId");
      if (!id) continue;
      if (attr(node, "type") === "table") {
        tableStyles.set(id, {
          basedOn: attr(first(node, "basedOn"), "val"),
          properties: tableStyleProperties(node),
        });
        if (attr(node, "default") === "1") defaultTableStyle = id;
      }
      styles.set(id, {
        basedOn: attr(first(node, "basedOn"), "val"),
        paragraph: paragraphProperties(first(node, "pPr")) as Props,
        run: runProperties(first(node, "rPr"), theme),
        numbering: (() => {
          const numPr = first(first(node, "pPr") ?? node, "numPr");
          const id = attr(first(numPr ?? node, "numId"), "val");
          const level = attr(first(numPr ?? node, "ilvl"), "val");
          return {
            numId: id,
            level: level === undefined ? undefined : Number(level),
          };
        })(),
      });
    }
  }

  function chain(id?: string): Style[] {
    const result: Style[] = [];
    const seen = new Set<string>();
    let current = id;
    while (current && !seen.has(current)) {
      seen.add(current);
      const style = styles.get(current);
      if (!style) break;
      result.unshift(style);
      current = style.basedOn;
    }
    return result;
  }

  function tableChain(id?: string): TableStyleProperties {
    const result: TableStyleProperties = { borders: {}, margins: {} };
    const seen = new Set<string>();
    let current = id ?? defaultTableStyle;
    while (current && !seen.has(current)) {
      seen.add(current);
      const style = tableStyles.get(current);
      if (!style) break;
      result.borders = { ...style.properties.borders, ...result.borders };
      result.margins = { ...style.properties.margins, ...result.margins };
      result.cellShading ??= style.properties.cellShading;
      current = style.basedOn;
    }
    return result;
  }

  return {
    paragraph(styleId, direct = {}) {
      const inherited = chain(styleId).map((item) => item.paragraph);
      const resolved = merge(defaultParagraph, ...inherited, direct as Props);
      return resolved as ParagraphProperties;
    },
    run(paragraphStyle, runStyle, direct = {}) {
      const pRun = chain(paragraphStyle).map((item) => item.run);
      const rRun = chain(runStyle).map((item) => item.run);
      return merge<TextStyle>(defaultRun, ...pRun, ...rRun, direct as Props);
    },
    numbering(styleId) {
      const resolved: { numId?: string; level?: number } = {};
      for (const style of chain(styleId)) {
        if (style.numbering.numId !== undefined) resolved.numId = style.numbering.numId;
        if (style.numbering.level !== undefined) resolved.level = style.numbering.level;
      }
      return resolved;
    },
    table(styleId) {
      return tableChain(styleId);
    },
  };
}

export function parseRunProperties(parent: Element | undefined, theme: Theme = {}): TextStyle {
  return runProperties(parent, theme) as TextStyle;
}
