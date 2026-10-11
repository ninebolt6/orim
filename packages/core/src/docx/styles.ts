import { attr, boolValue, descendants, first, value as xmlValue } from "../xml";
import { border } from "./borders";
import { halfPoints, points, twips } from "./units";
import type { ThemeFonts } from "./theme";
import type { ParagraphProperties, TextStyle } from "./model";

type BorderSides = Partial<
  Record<"top" | "right" | "bottom" | "left" | "insideH" | "insideV", string>
>;
type CellMargins = Partial<Record<"top" | "right" | "bottom" | "left", number>>;
export type RawRunProperties = Partial<TextStyle> & {
  fontLatin?: string;
  fontEastAsia?: string;
};

export interface TableStyleProperties {
  borders: BorderSides;
  margins: CellMargins;
  cellShading?: string;
}

interface Style {
  basedOn?: string;
  paragraph: Partial<ParagraphProperties>;
  run: RawRunProperties;
  numbering: { numId?: string; level?: number };
}

export interface StyleSheet {
  paragraph(styleId?: string, direct?: Partial<ParagraphProperties>): ParagraphProperties;
  run(paragraphStyle?: string, runStyle?: string, direct?: RawRunProperties): TextStyle;
  numbering(styleId?: string): { numId?: string; level?: number };
  table(styleId?: string): TableStyleProperties;
}

const safeColor = (val: string | undefined): string | undefined => {
  return val && /^[\da-f]{6}$/i.test(val) ? `#${val}` : undefined;
};

const tableStyleProperties = (node: Element): TableStyleProperties => {
  const tableProperties = first(node, "tblPr");
  const borderProperties = first(tableProperties, "tblBorders");
  const borders = Object.fromEntries(
    (["top", "right", "bottom", "left", "insideH", "insideV"] as const).flatMap((side) => {
      const value = border(first(borderProperties, side));
      return value === undefined ? [] : [[side, value] as const];
    }),
  );
  const marginProperties = first(tableProperties, "tblCellMar");
  const margins = Object.fromEntries(
    (["top", "right", "bottom", "left"] as const).flatMap((side) => {
      const amount = twips(attr(first(marginProperties, side), "w"));
      return amount === undefined ? [] : [[side, amount] as const];
    }),
  );
  const shading = first(first(node, "tcPr"), "shd");
  const fill = attr(shading, "fill");
  return { borders, margins, cellShading: fill && fill !== "auto" ? safeColor(fill) : undefined };
};

type Theme = ThemeFonts;

const runProperties = (parent: Element | undefined, theme: Theme): RawRunProperties => {
  if (!parent) return {};
  const fonts = first(parent, "rFonts");
  const latinTheme = attr(fonts, "asciiTheme") ?? attr(fonts, "hAnsiTheme");
  const eastAsiaTheme = attr(fonts, "eastAsiaTheme");
  const size = halfPoints(xmlValue(first(parent, "sz")));
  const color = xmlValue(first(parent, "color"));
  const highlight = xmlValue(first(parent, "highlight"));
  const vertAlign = xmlValue(first(parent, "vertAlign"));
  const underline = xmlValue(first(parent, "u"));
  const shading = first(parent, "shd");
  const fill = attr(shading, "fill");
  return {
    bold: boolValue(first(parent, "b")),
    italic: boolValue(first(parent, "i")),
    underline:
      underline === "none" ? undefined : (underline ?? (first(parent, "u") ? "single" : undefined)),
    strike: boolValue(first(parent, "strike")),
    doubleStrike: boolValue(first(parent, "dstrike")),
    color: color === "auto" ? undefined : safeColor(color),
    highlight: highlight && highlight !== "none" ? highlight : undefined,
    fontLatin:
      attr(fonts, "ascii") ??
      attr(fonts, "hAnsi") ??
      (latinTheme ? (latinTheme.startsWith("major") ? theme.major : theme.minor) : undefined),
    fontEastAsia:
      attr(fonts, "eastAsia") ??
      (eastAsiaTheme
        ? eastAsiaTheme.startsWith("major")
          ? theme.majorEastAsia
          : theme.minorEastAsia
        : undefined),
    fontSize: size,
    verticalAlign: vertAlign === "superscript" || vertAlign === "subscript" ? vertAlign : undefined,
    caps: boolValue(first(parent, "caps")),
    smallCaps: boolValue(first(parent, "smallCaps")),
    hidden: boolValue(first(parent, "vanish")),
    ...(fill && fill !== "auto" ? { backgroundColor: safeColor(fill) } : {}),
  };
};

const composeRunProperties = (props: RawRunProperties, theme: Theme): TextStyle => {
  const { fontLatin, fontEastAsia, ...style } = props;
  const explicit = [fontLatin, fontEastAsia].filter(
    (item, index, all): item is string => !!item && all.indexOf(item) === index,
  );
  const fontFamily = explicit.length > 0 ? explicit.join(",") : (style.fontFamily ?? theme.minor);
  return { ...style, fontFamily };
};

export const paragraphProperties = (parent: Element | undefined): Partial<ParagraphProperties> => {
  if (!parent) return {};
  const spacingNode = first(parent, "spacing");
  const indent = first(parent, "ind");
  const jc = xmlValue(first(parent, "jc"));
  const shading = first(parent, "shd");
  const fill = attr(shading, "fill");
  const line = attr(spacingNode, "line");
  const lineRule = attr(spacingNode, "lineRule") ?? (line === undefined ? undefined : "auto");
  const base: Partial<ParagraphProperties> = {
    align: jc,
    before: twips(attr(spacingNode, "before")),
    after: twips(attr(spacingNode, "after")),
    line: lineRule === "auto" ? points(line, 240) : twips(line),
    lineRule,
    left: twips(attr(indent, "left") ?? attr(indent, "start")),
    right: twips(attr(indent, "right") ?? attr(indent, "end")),
    firstLine: twips(attr(indent, "firstLine")),
    hanging: twips(attr(indent, "hanging")),
    contextualSpacing: boolValue(first(parent, "contextualSpacing")),
    pageBreakBefore: boolValue(first(parent, "pageBreakBefore")),
    keepNext: boolValue(first(parent, "keepNext")),
    snapToGrid: boolValue(first(parent, "snapToGrid")),
    shading: fill && fill !== "auto" ? safeColor(fill) : undefined,
  };
  const borders = first(parent, "pBdr");
  const sides = ["top", "right", "bottom", "left"] as const;
  const values = Object.fromEntries(
    sides.flatMap((side) => {
      const value = border(first(borders, side));
      return value === undefined ? [] : [[side, value] as const];
    }),
  );
  const spacing = Object.fromEntries(
    sides.flatMap((side) => {
      const amount = points(attr(first(borders, side), "space"), 1);
      return amount === undefined ? [] : [[side, amount] as const];
    }),
  );
  return {
    ...base,
    ...(Object.keys(values).length ? { borders: values } : {}),
    ...(Object.keys(spacing).length ? { borderSpacing: spacing } : {}),
  };
};

const isObject = (value: unknown): value is object => {
  return typeof value === "object" && value !== null;
};

export const mergeLayers = <T extends object>(
  layers: readonly (Partial<T> | undefined)[],
  nestedKeys: readonly (keyof T)[],
): Partial<T> => {
  const merged: { -readonly [Key in keyof T]?: T[Key] } = {};
  for (const layer of layers) {
    if (!layer) continue;
    for (const key of Object.keys(layer) as Array<keyof T>) {
      const value = layer[key];
      if (value === undefined) continue;
      const previous = merged[key];
      const next =
        nestedKeys.includes(key) && isObject(value)
          ? { ...(isObject(previous) ? previous : {}), ...value }
          : value;
      Reflect.set(merged, key, next);
    }
  }
  return merged;
};

export const parseStyles = (doc: Document | undefined, theme: Theme = {}): StyleSheet => {
  const styles = new Map<string, Style>();
  const tableStyles = new Map<string, { basedOn?: string; properties: TableStyleProperties }>();
  let defaultTableStyle: string | undefined;
  const defaults = doc && descendants(doc, "docDefaults")[0];
  const defaultParagraph = paragraphProperties(first(first(defaults, "pPrDefault"), "pPr"));
  const defaultRun = runProperties(first(first(defaults, "rPrDefault"), "rPr"), theme);
  if (doc) {
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
        paragraph: paragraphProperties(first(node, "pPr")),
        run: runProperties(first(node, "rPr"), theme),
        numbering: (() => {
          const numPr = first(first(node, "pPr"), "numPr");
          const id = attr(first(numPr, "numId"), "val");
          const level = attr(first(numPr, "ilvl"), "val");
          return {
            numId: id,
            level: level === undefined ? undefined : Number(level),
          };
        })(),
      });
    }
  }

  const chainCache = new Map<string, readonly Style[]>();
  const chain = (id?: string): readonly Style[] => {
    const key = id ?? "";
    const cached = chainCache.get(key);
    if (cached) return cached;
    const reversed: Style[] = [];
    const visited = new Set<string>();
    let current = id;
    while (current && !visited.has(current)) {
      visited.add(current);
      const style = styles.get(current);
      if (!style) break;
      reversed.push(style);
      current = style.basedOn;
    }
    const result = reversed.reverse();
    chainCache.set(key, result);
    return result;
  };

  const tableChain = (id?: string): TableStyleProperties => {
    const resolved: TableStyleProperties = { borders: {}, margins: {} };
    const visited = new Set<string>();
    let current = id ?? defaultTableStyle;
    while (current && !visited.has(current)) {
      visited.add(current);
      const style = tableStyles.get(current);
      if (!style) break;
      for (const [side, value] of Object.entries(style.properties.borders))
        if (resolved.borders[side as keyof BorderSides] === undefined)
          Reflect.set(resolved.borders, side, value);
      for (const [side, value] of Object.entries(style.properties.margins))
        if (resolved.margins[side as keyof CellMargins] === undefined)
          Reflect.set(resolved.margins, side, value);
      resolved.cellShading ??= style.properties.cellShading;
      current = style.basedOn;
    }
    return resolved;
  };

  const paragraphCache = new Map<string, Partial<ParagraphProperties>>();
  const runCache = new Map<string, RawRunProperties>();
  const inheritedParagraph = (styleId?: string) => {
    const key = styleId ?? "";
    const cached = paragraphCache.get(key);
    if (cached) return cached;
    const result = mergeLayers<ParagraphProperties>(
      [defaultParagraph, ...chain(styleId).map((item) => item.paragraph)],
      ["borders", "borderSpacing"],
    );
    paragraphCache.set(key, result);
    return result;
  };
  const inheritedRun = (paragraphStyle?: string, runStyle?: string) => {
    const key = JSON.stringify([paragraphStyle ?? "", runStyle ?? ""]);
    const cached = runCache.get(key);
    if (cached) return cached;
    const result = mergeLayers<RawRunProperties>(
      [
        defaultRun,
        ...chain(paragraphStyle).map((item) => item.run),
        ...chain(runStyle).map((item) => item.run),
      ],
      [],
    );
    runCache.set(key, result);
    return result;
  };

  return {
    paragraph(styleId, direct = {}) {
      return mergeLayers<ParagraphProperties>(
        [inheritedParagraph(styleId), direct],
        ["borders", "borderSpacing"],
      );
    },
    run(paragraphStyle, runStyle, direct = {}) {
      const inherited = inheritedRun(paragraphStyle, runStyle);
      const hasDirectScriptFont =
        direct.fontLatin !== undefined || direct.fontEastAsia !== undefined;
      const cascade =
        direct.fontFamily !== undefined && !hasDirectScriptFont
          ? { ...inherited, fontLatin: undefined, fontEastAsia: undefined }
          : inherited;
      return composeRunProperties(mergeLayers<RawRunProperties>([cascade, direct], []), theme);
    },
    numbering(styleId) {
      const resolved: { numId?: string; level?: number } = {};
      for (const style of chain(styleId)) {
        resolved.numId = style.numbering.numId ?? resolved.numId;
        resolved.level = style.numbering.level ?? resolved.level;
      }
      return resolved;
    },
    table(styleId) {
      return tableChain(styleId);
    },
  };
};

export const parseRunProperties = (
  parent: Element | undefined,
  theme: Theme = {},
): RawRunProperties => {
  return runProperties(parent, theme);
};
