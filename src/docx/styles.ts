import { attr, descendants, elements, first } from "../opc";
import type { ParagraphProperties, TextStyle } from "./model";

type Props = Record<string, string | number | boolean | undefined>;

interface Style {
  basedOn?: string;
  paragraph: Props;
  run: Props;
}

export interface StyleSheet {
  paragraph(styleId?: string, direct?: Partial<ParagraphProperties>): ParagraphProperties;
  run(paragraphStyle?: string, runStyle?: string, direct?: Partial<TextStyle>): TextStyle;
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

function paragraphProperties(parent: Element | undefined): Props {
  if (!parent) return {};
  const spacing = first(parent, "spacing");
  const indent = first(parent, "ind");
  const jc = value(first(parent, "jc"));
  const shading = first(parent, "shd");
  const fill = attr(shading, "fill");
  const lineRule = attr(spacing, "lineRule");
  const line = attr(spacing, "line");
  const props: Props = {
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
    shading: fill && fill !== "auto" ? safeColor(fill) : undefined,
  };
  const borders = first(parent, "pBdr");
  if (borders) {
    const side = elements(borders, "left")[0] ?? elements(borders, "bottom")[0];
    if (side) {
      const color = safeColor(attr(side, "color"));
      const size = Number(attr(side, "sz") ?? 4) / 8;
      props.border = `${size}pt solid ${color ?? "#888"}`;
    }
  }
  return props;
}

function merge<T>(...layers: Array<Props | undefined>): T {
  const out: Props = {};
  for (const layer of layers) {
    if (!layer) continue;
    for (const [key, item] of Object.entries(layer)) if (item !== undefined) out[key] = item;
  }
  return out as T;
}

export function parseStyles(doc: Document | undefined, theme: Theme = {}): StyleSheet {
  const styles = new Map<string, Style>();
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
      styles.set(id, {
        basedOn: attr(first(node, "basedOn"), "val"),
        paragraph: paragraphProperties(first(node, "pPr")),
        run: runProperties(first(node, "rPr"), theme),
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
  };
}

export function parseRunProperties(parent: Element | undefined, theme: Theme = {}): TextStyle {
  return runProperties(parent, theme) as TextStyle;
}
