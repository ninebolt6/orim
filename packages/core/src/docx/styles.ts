import { attr, boolValue, descendants, first, value as xmlValue } from "./xml";
import { border } from "./borders";
import { halfPoints, points, twips } from "./units";
import { themeEastAsiaFont, type ThemeFonts } from "../drawingml/theme";
import { safeListLevel } from "./numbering";
import type { ParagraphProperties, TextStyle } from "./model";

type BorderSides = Partial<
  Record<"top" | "right" | "bottom" | "left" | "insideH" | "insideV", string>
>;
type CellMargins = Partial<Record<"top" | "right" | "bottom" | "left", number>>;
export type RawRunProperties = Partial<TextStyle> & {
  fontLatin?: string;
  fontEastAsia?: string;
  fontEastAsiaLanguage?: string;
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
  run(
    paragraphStyle?: string,
    runStyle?: string,
    direct?: RawRunProperties,
    text?: string,
  ): TextStyle;
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
    underline: underline ?? (first(parent, "u") ? "single" : undefined),
    strike: boolValue(first(parent, "strike")),
    doubleStrike: boolValue(first(parent, "dstrike")),
    color: color === "auto" ? undefined : safeColor(color),
    highlight: highlight && highlight !== "none" ? highlight : undefined,
    fontLatin:
      attr(fonts, "ascii") ??
      attr(fonts, "hAnsi") ??
      (latinTheme ? (latinTheme.startsWith("major") ? theme.major : theme.minor) : undefined),
    fontEastAsia: attr(fonts, "eastAsia") ?? eastAsiaTheme,
    fontEastAsiaLanguage: attr(first(parent, "lang"), "eastAsia"),
    fontSize: size,
    verticalAlign: vertAlign === "superscript" || vertAlign === "subscript" ? vertAlign : undefined,
    caps: boolValue(first(parent, "caps")),
    smallCaps: boolValue(first(parent, "smallCaps")),
    hidden: boolValue(first(parent, "vanish")),
    ...(fill && fill !== "auto" ? { backgroundColor: safeColor(fill) } : {}),
  };
};

const themeFontReferences = new Set([
  "majorAscii",
  "majorHAnsi",
  "majorEastAsia",
  "majorBidi",
  "minorAscii",
  "minorHAnsi",
  "minorEastAsia",
  "minorBidi",
]);

const themeFontReference = (
  reference: string,
  theme: Theme,
  language: string | undefined,
  text: string,
): string | undefined => {
  switch (reference) {
    case "majorAscii":
    case "majorHAnsi":
      return theme.major;
    case "minorAscii":
    case "minorHAnsi":
      return theme.minor;
    case "majorEastAsia":
      return themeEastAsiaFont(theme, true, language, text);
    case "minorEastAsia":
      return themeEastAsiaFont(theme, false, language, text);
    case "majorBidi":
      return theme.majorCs;
    case "minorBidi":
      return theme.minorCs;
  }
  return undefined;
};

const composeRunProperties = (props: RawRunProperties, theme: Theme, text = ""): TextStyle => {
  const { fontLatin, fontEastAsia, fontEastAsiaLanguage, ...style } = props;
  const resolvedEastAsia = fontEastAsia
    ? themeFontReferences.has(fontEastAsia)
      ? themeFontReference(fontEastAsia, theme, fontEastAsiaLanguage, text)
      : fontEastAsia
    : undefined;
  const explicit = [fontLatin, resolvedEastAsia].filter(
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

const breakBasedOnCycles = <T extends { basedOn?: string }>(items: Map<string, T>): void => {
  const state = new Map<string, 1 | 2>();
  for (const id of items.keys()) {
    if (state.has(id)) continue;
    const path: string[] = [];
    let current: string | undefined = id;
    while (current) {
      const item = items.get(current);
      if (!item) break;
      const itemState = state.get(current);
      if (itemState === 1) {
        const cycleStart = path.indexOf(current);
        if (cycleStart >= 0) {
          const cycle = path.slice(cycleStart);
          const cut = cycle.reduce((larger, value) => (value > larger ? value : larger));
          const target = items.get(cut);
          if (target) target.basedOn = undefined;
        }
        break;
      }
      if (itemState === 2) break;
      state.set(current, 1);
      path.push(current);
      current = item.basedOn;
    }
    for (const node of path) state.set(node, 2);
  }
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
            level: level === undefined ? undefined : safeListLevel(level),
          };
        })(),
      });
    }
  }
  breakBasedOnCycles(styles);
  breakBasedOnCycles(tableStyles);

  const chainValues = <S, T extends object>(
    lookup: (id: string) => S | undefined,
    basedOn: (style: S) => string | undefined,
    properties: (style: S) => Partial<T>,
    nestedKeys: readonly (keyof T)[],
  ): ((id?: string) => Partial<T>) => {
    const cache = new Map<string, Partial<T>>();
    return (id?: string) => {
      const key = id ?? "";
      const cached = cache.get(key);
      if (cached) return cached;
      const path: string[] = [];
      const seen = new Set<string>();
      let current = id;
      let merged: Partial<T> = {};
      while (current) {
        const cachedValue = cache.get(current);
        if (cachedValue) {
          merged = cachedValue;
          break;
        }
        if (seen.has(current)) break;
        seen.add(current);
        const style = lookup(current);
        if (!style) break;
        path.push(current);
        current = basedOn(style);
      }
      for (let index = path.length - 1; index >= 0; index--) {
        const styleId = path[index];
        const style = lookup(styleId);
        if (style) merged = mergeLayers<T>([merged, properties(style)], nestedKeys);
        cache.set(styleId, merged);
      }
      cache.set(key, merged);
      return merged;
    };
  };

  const styleChain = <T extends object>(
    properties: (style: Style) => Partial<T>,
    nestedKeys: readonly (keyof T)[],
  ) =>
    chainValues<Style, T>(
      (id) => styles.get(id),
      (style) => style.basedOn,
      properties,
      nestedKeys,
    );
  const paragraphChain = styleChain<ParagraphProperties>(
    (style) => style.paragraph,
    ["borders", "borderSpacing"],
  );
  const runChain = styleChain<RawRunProperties>((style) => style.run, []);
  const numberingChain = styleChain<{ numId?: string; level?: number }>(
    (style) => style.numbering,
    [],
  );
  const tableChainValues = chainValues<
    { basedOn?: string; properties: TableStyleProperties },
    TableStyleProperties
  >(
    (id) => tableStyles.get(id),
    (style) => style.basedOn,
    (style) => style.properties,
    ["borders", "margins"],
  );

  const tableCache = new Map<string | undefined, TableStyleProperties>();
  const tableChain = (id?: string): TableStyleProperties => {
    const cached = tableCache.get(id);
    if (cached) return cached;
    const inherited = tableChainValues(id ?? defaultTableStyle);
    const result: TableStyleProperties = {
      borders: inherited.borders ?? {},
      margins: inherited.margins ?? {},
      ...(inherited.cellShading ? { cellShading: inherited.cellShading } : {}),
    };
    tableCache.set(id, result);
    return result;
  };

  const paragraphCache = new Map<string, Partial<ParagraphProperties>>();
  const runCache = new Map<string, RawRunProperties>();
  const inheritedParagraph = (styleId?: string) => {
    const key = styleId ?? "";
    const cached = paragraphCache.get(key);
    if (cached) return cached;
    const result = mergeLayers<ParagraphProperties>(
      [defaultParagraph, paragraphChain(styleId)],
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
      [defaultRun, runChain(paragraphStyle), runChain(runStyle)],
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
    run(paragraphStyle, runStyle, direct = {}, text = "") {
      const inherited = inheritedRun(paragraphStyle, runStyle);
      const hasDirectScriptFont =
        direct.fontLatin !== undefined || direct.fontEastAsia !== undefined;
      const cascade =
        direct.fontFamily !== undefined && !hasDirectScriptFont
          ? { ...inherited, fontLatin: undefined, fontEastAsia: undefined }
          : inherited;
      return composeRunProperties(
        mergeLayers<RawRunProperties>([cascade, direct], []),
        theme,
        text,
      );
    },
    numbering(styleId) {
      return numberingChain(styleId);
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
