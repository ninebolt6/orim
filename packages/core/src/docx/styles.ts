import { attr, boolValue, descendants, first, value as xmlValue } from "../xml";
import { border } from "./borders";
import { halfPoints, points, twips } from "./units";
import type { ParagraphProperties, TextStyle } from "./model";

type Props = Record<string, unknown>;
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

function safeColor(val: string | undefined): string | undefined {
  return val && /^[\da-f]{6}$/i.test(val) ? `#${val}` : undefined;
}

function tableStyleProperties(node: Element): TableStyleProperties {
  const tableProperties = first(node, "tblPr");
  const borderProperties = first(tableProperties ?? node, "tblBorders");
  const borders = Object.fromEntries(
    (["top", "right", "bottom", "left", "insideH", "insideV"] as const).flatMap((side) => {
      const value = border(first(borderProperties ?? node, side));
      return value === undefined ? [] : [[side, value] as const];
    }),
  );
  const marginProperties = first(tableProperties ?? node, "tblCellMar");
  const margins = Object.fromEntries(
    (["top", "right", "bottom", "left"] as const).flatMap((side) => {
      const amount = twips(attr(first(marginProperties ?? node, side), "w"));
      return amount === undefined ? [] : [[side, amount] as const];
    }),
  );
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
    fontFamily: family || undefined,
    fontSize: size,
    verticalAlign: vertAlign === "superscript" || vertAlign === "subscript" ? vertAlign : undefined,
    caps: boolValue(first(parent, "caps")),
    smallCaps: boolValue(first(parent, "smallCaps")),
    hidden: boolValue(first(parent, "vanish")),
    ...(fill && fill !== "auto" ? { backgroundColor: safeColor(fill) } : {}),
  };
}

export function paragraphProperties(parent: Element | undefined): Partial<ParagraphProperties> {
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
}

function merge<T>(...layers: Array<Props | undefined>): T {
  return layers
    .filter((layer): layer is Props => !!layer)
    .reduce<Props>(
      (merged, layer) =>
        Object.entries(layer).reduce<Props>((next, [key, value]) => {
          if (value === undefined) return next;
          const previous = next[key];
          return {
            ...next,
            [key]:
              (key === "borders" || key === "borderSpacing") && value && typeof value === "object"
                ? {
                    ...(previous && typeof previous === "object" ? previous : {}),
                    ...value,
                  }
                : value,
          };
        }, merged),
      {},
    ) as T;
}

export function parseStyles(doc: Document | undefined, theme: Theme = {}): StyleSheet {
  const styles = new Map<string, Style>();
  const tableStyles = new Map<string, { basedOn?: string; properties: TableStyleProperties }>();
  let defaultTableStyle: string | undefined;
  const defaults = doc && descendants(doc, "docDefaults")[0];
  const defaultParagraph: Props = doc
    ? (paragraphProperties(first(first(defaults ?? doc, "pPrDefault") ?? doc, "pPr")) as Props)
    : {};
  const defaultRun: Props = doc
    ? runProperties(first(first(defaults ?? doc, "rPrDefault") ?? doc, "rPr"), theme)
    : {};
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
        paragraph: paragraphProperties(first(node, "pPr")) as unknown as Props,
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

  function chain(id?: string, visited: readonly string[] = []): Style[] {
    if (!id || visited.includes(id)) return [];
    const style = styles.get(id);
    return style ? [...chain(style.basedOn, [...visited, id]), style] : [];
  }

  function tableChain(id?: string): TableStyleProperties {
    const resolve = (current?: string, visited: readonly string[] = []): TableStyleProperties[] => {
      if (!current || visited.includes(current)) return [];
      const style = tableStyles.get(current);
      return style ? [style.properties, ...resolve(style.basedOn, [...visited, current])] : [];
    };
    return resolve(id ?? defaultTableStyle).reduce<TableStyleProperties>(
      (resolved, properties) => {
        return {
          borders: { ...properties.borders, ...resolved.borders },
          margins: { ...properties.margins, ...resolved.margins },
          cellShading: resolved.cellShading ?? properties.cellShading,
        };
      },
      { borders: {}, margins: {} },
    );
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
      return chain(styleId).reduce(
        (resolved, style) => ({
          numId: style.numbering.numId ?? resolved.numId,
          level: style.numbering.level ?? resolved.level,
        }),
        {} as { numId?: string; level?: number },
      );
    },
    table(styleId) {
      return tableChain(styleId);
    },
  };
}

export function parseRunProperties(parent: Element | undefined, theme: Theme = {}): TextStyle {
  return runProperties(parent, theme) as TextStyle;
}
