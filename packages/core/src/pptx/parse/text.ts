import { safeHref } from "../../common/sanitize";
import { alphabetic, roman } from "../../common/numbering";
import { attr, children, elements, first, R } from "../../common/xml";
import { parseColor } from "../../drawingml/color";
import { emu, fraction } from "../../drawingml/units";
import type {
  Bullet,
  Insets,
  LineBreak,
  RunStyle,
  Spacing,
  TextBody,
  TextParagraph,
  TextRun,
} from "../model";
import { defined, layer, type ShapeContext, type SlideContext } from "./context";

type BulletSize = { readonly scale: number } | { readonly points: number };

export interface ParagraphProps {
  readonly align?: TextParagraph["align"];
  readonly marginLeft?: number;
  readonly indent?: number;
  readonly lineSpacing?: Spacing;
  readonly spaceBefore?: Spacing;
  readonly spaceAfter?: Spacing;
  readonly bullet?: "none" | "char" | "auto";
  readonly bulletChar?: string;
  readonly bulletScheme?: string;
  readonly bulletStart?: number;
  /** Bullet typeface; an empty string follows the text. */
  readonly bulletFont?: string;
  /** Bullet color; an empty string follows the text. */
  readonly bulletColor?: string;
  readonly bulletSize?: BulletSize;
  readonly run?: Partial<RunStyle>;
}

/** Paragraph defaults for outline levels 1–9. */
export type ListStyle = readonly ParagraphProps[];

export interface BodyProps {
  readonly anchor?: TextBody["anchor"];
  readonly left?: number;
  readonly top?: number;
  readonly right?: number;
  readonly bottom?: number;
  readonly wrap?: boolean;
  readonly vertical?: TextBody["vertical"] | "horizontal";
  readonly fontScale?: number;
  readonly lineReduction?: number;
}

const levels = Array.from({ length: 9 }, (_, index) => index);

const flag = (value: string | undefined): boolean | undefined =>
  value === undefined ? undefined : value === "1" || value === "true";

const hundredths = (value: string | undefined): number | undefined => {
  const parsed = Number(value);
  return value !== undefined && Number.isFinite(parsed) ? parsed / 100 : undefined;
};

const typeface = (element: Element | undefined, context: SlideContext): string | undefined => {
  const value = attr(element, "typeface");
  switch (value) {
    case undefined:
    case "":
      return undefined;
    case "+mj-lt":
      return context.fonts.major;
    case "+mn-lt":
      return context.fonts.minor;
    case "+mj-ea":
      return context.fonts.majorEastAsia;
    case "+mn-ea":
      return context.fonts.minorEastAsia;
  }
  return value.startsWith("+") ? undefined : value;
};

const textColor = (properties: Element, context: SlideContext): string | undefined => {
  const solid = first(properties, "solidFill");
  if (solid) return parseColor(solid, context.scheme);
  const stop = first(first(first(properties, "gradFill"), "gsLst"), "gs");
  return stop && parseColor(stop, context.scheme);
};

export const parseRunProps = (
  properties: Element | undefined,
  context: SlideContext,
): Partial<RunStyle> => {
  if (!properties) return {};
  const underline = attr(properties, "u");
  const strike = attr(properties, "strike");
  const caps = attr(properties, "cap");
  return defined({
    size: hundredths(attr(properties, "sz")),
    bold: flag(attr(properties, "b")),
    italic: flag(attr(properties, "i")),
    underline: underline === undefined ? undefined : underline !== "none",
    strike: strike === undefined ? undefined : strike !== "noStrike",
    color: textColor(properties, context),
    highlight: parseColor(first(properties, "highlight"), context.scheme),
    font: typeface(first(properties, "latin"), context),
    eastAsianFont: typeface(first(properties, "ea"), context),
    baseline: fraction(attr(properties, "baseline")),
    caps: caps === "all" || caps === "small" ? caps : undefined,
    spacing: hundredths(attr(properties, "spc")),
  });
};

const spacing = (element: Element | undefined): Spacing | undefined => {
  const points = hundredths(attr(first(element, "spcPts"), "val"));
  if (points !== undefined) return { points };
  const percent = fraction(attr(first(element, "spcPct"), "val"));
  return percent === undefined ? undefined : { percent };
};

const alignments: Readonly<Record<string, TextParagraph["align"]>> = {
  l: "left",
  ctr: "center",
  r: "right",
  just: "justify",
  dist: "justify",
};

export const parseParagraphProps = (
  properties: Element | undefined,
  context: SlideContext,
): ParagraphProps => {
  if (!properties) return {};
  const has = (name: string) => !!first(properties, name);
  const autoNumber = first(properties, "buAutoNum");
  const sizePercent = fraction(attr(first(properties, "buSzPct"), "val"));
  const sizePoints = hundredths(attr(first(properties, "buSzPts"), "val"));
  const start = Number(attr(autoNumber, "startAt"));
  return defined({
    align: alignments[attr(properties, "algn") ?? ""],
    marginLeft: emu(attr(properties, "marL")),
    indent: emu(attr(properties, "indent")),
    lineSpacing: spacing(first(properties, "lnSpc")),
    spaceBefore: spacing(first(properties, "spcBef")),
    spaceAfter: spacing(first(properties, "spcAft")),
    bullet: has("buNone")
      ? "none"
      : has("buChar") || has("buBlip")
        ? "char"
        : autoNumber
          ? "auto"
          : undefined,
    bulletChar: has("buBlip") ? "•" : attr(first(properties, "buChar"), "char"),
    bulletScheme: attr(autoNumber, "type"),
    bulletStart: autoNumber ? (Number.isInteger(start) && start > 0 ? start : 1) : undefined,
    bulletFont: has("buFontTx") ? "" : typeface(first(properties, "buFont"), context),
    bulletColor: has("buClrTx") ? "" : parseColor(first(properties, "buClr"), context.scheme),
    bulletSize: has("buSzTx")
      ? { scale: 1 }
      : sizePercent !== undefined
        ? { scale: sizePercent }
        : sizePoints !== undefined
          ? { points: sizePoints }
          : undefined,
    run: first(properties, "defRPr") && parseRunProps(first(properties, "defRPr"), context),
  });
};

const mergeParagraph = (
  base: ParagraphProps | undefined,
  override: ParagraphProps | undefined,
): ParagraphProps => ({
  ...layer<ParagraphProps>(base, override),
  run: layer<RunStyle>(base?.run, override?.run),
});

export const parseListStyle = (element: Element | undefined, context: SlideContext): ListStyle => {
  const defaults = parseParagraphProps(first(element, "defPPr"), context);
  return levels.map((level) =>
    mergeParagraph(defaults, parseParagraphProps(first(element, `lvl${level + 1}pPr`), context)),
  );
};

export const uniformListStyle = (run: Partial<RunStyle>): ListStyle => levels.map(() => ({ run }));

export const mergeListStyles = (...styles: readonly (ListStyle | undefined)[]): ListStyle =>
  levels.map((level) =>
    styles.reduce<ParagraphProps>((result, style) => mergeParagraph(result, style?.[level]), {}),
  );

export const parseBodyProps = (element: Element | undefined): BodyProps => {
  if (!element) return {};
  const anchor = attr(element, "anchor");
  const vertical = attr(element, "vert");
  const normal = first(element, "normAutofit");
  const fitted = normal ?? first(element, "noAutofit") ?? first(element, "spAutoFit");
  return defined({
    anchor: anchor === "ctr" ? "middle" : anchor === "b" ? "bottom" : anchor ? "top" : undefined,
    left: emu(attr(element, "lIns")),
    top: emu(attr(element, "tIns")),
    right: emu(attr(element, "rIns")),
    bottom: emu(attr(element, "bIns")),
    wrap: attr(element, "wrap") === undefined ? undefined : attr(element, "wrap") !== "none",
    vertical:
      vertical === undefined
        ? undefined
        : vertical === "vert" || vertical === "eaVert" || vertical === "mongolianVert"
          ? "vertical"
          : vertical === "vert270"
            ? "vertical270"
            : vertical === "wordArtVert" || vertical === "wordArtVertRtl"
              ? "stacked"
              : "horizontal",
    fontScale: fitted ? (fraction(attr(normal, "fontScale")) ?? 1) : undefined,
    lineReduction: fitted ? (fraction(attr(normal, "lnSpcReduction")) ?? 0) : undefined,
  });
};

const symbolBullets: Readonly<Record<string, string>> = {
  "§": "▪",
  Ø: "➢",
  ü: "✓",
  q: "❑",
  v: "❖",
  n: "■",
  l: "●",
  u: "◆",
  Ÿ: "•",
  "·": "•",
  à: "→",
  è: "➔",
};

const numberText = (scheme: string | undefined, number: number): string => {
  const match =
    /^(arabic|alphaLc|alphaUc|romanLc|romanUc)(Period|ParenR|ParenBoth|Plain|Minus)?/.exec(
      scheme ?? "",
    );
  const [, style = "arabic", suffix = "Period"] = match ?? [];
  const value =
    style === "alphaLc"
      ? alphabetic(number, false)
      : style === "alphaUc"
        ? alphabetic(number, true)
        : style === "romanLc"
          ? roman(number).toLowerCase()
          : style === "romanUc"
            ? roman(number)
            : String(number);
  return suffix === "ParenR"
    ? `${value})`
    : suffix === "ParenBoth"
      ? `(${value})`
      : suffix === "Plain"
        ? value
        : suffix === "Minus"
          ? `- ${value} -`
          : `${value}.`;
};

const isSymbolFont = (font: string | undefined): boolean =>
  !!font && /^(?:wingdings|symbol)/i.test(font);

const bulletFor = (
  props: ParagraphProps,
  number: number,
  firstStyle: RunStyle,
): Bullet | undefined => {
  if (props.bullet !== "char" && props.bullet !== "auto") return undefined;
  const font = props.bulletFont || undefined;
  const raw = props.bullet === "auto" ? numberText(props.bulletScheme, number) : props.bulletChar;
  if (!raw) return undefined;
  const symbol = props.bullet === "char" && isSymbolFont(font) ? symbolBullets[raw] : undefined;
  return defined({
    text: symbol ?? raw,
    font: symbol || isSymbolFont(font) ? undefined : font,
    color: props.bulletColor || firstStyle.color,
    scale: props.bulletSize && "scale" in props.bulletSize ? props.bulletSize.scale : undefined,
    size: props.bulletSize && "points" in props.bulletSize ? props.bulletSize.points : undefined,
  }) as Bullet;
};

export const linkTarget = (
  properties: Element | undefined,
  context: ShapeContext,
): string | undefined => {
  const id = attr(first(properties, "hlinkClick"), "id", R);
  const relation = id ? context.part.rels.get(id) : undefined;
  if (!relation) return undefined;
  if (relation.external) return safeHref(relation.target);
  const slide = context.slidePaths.get(
    context.part.opc.resolve(context.part.path, relation.target),
  );
  return slide === undefined ? undefined : `#slide-${slide}`;
};

const runs = (paragraph: Element, base: RunStyle, context: ShapeContext): (TextRun | LineBreak)[] =>
  children(paragraph).flatMap((child): (TextRun | LineBreak)[] => {
    const properties = first(child, "rPr");
    const style = { ...base, ...parseRunProps(properties, context) };
    if (child.localName === "br") return [{ type: "break", style }];
    if (child.localName !== "r" && child.localName !== "fld") return [];
    const text =
      child.localName === "fld" && attr(child, "type") === "slidenum"
        ? String(context.slideNumber)
        : elements(child, "t")
            .map((item) => item.textContent ?? "")
            .join("");
    if (!text) return [];
    const href = linkTarget(properties, context);
    return [
      href
        ? {
            text,
            style: { ...style, underline: true, color: context.scheme("hlink") ?? style.color },
            href,
          }
        : { text, style },
    ];
  });

const defaultInsets: Insets = { left: 7.2, top: 3.6, right: 7.2, bottom: 3.6 };

/**
 * Parses `p:txBody`/`a:txBody`. `styles` and `bodies` are inherited layers ordered from
 * lowest to highest priority; the body's own `a:lstStyle` and `a:bodyPr` are applied last.
 */
export const parseTextBody = (
  body: Element | undefined,
  styles: readonly (ListStyle | undefined)[],
  bodies: readonly BodyProps[],
  context: ShapeContext,
): TextBody | undefined => {
  if (!body) return undefined;
  const listStyle = mergeListStyles(...styles, parseListStyle(first(body, "lstStyle"), context));
  const bodyProps = layer<BodyProps>(...bodies, parseBodyProps(first(body, "bodyPr")));
  const counters: (number | undefined)[] = [];
  const paragraphs = elements(body, "p").map((paragraph): TextParagraph => {
    const properties = first(paragraph, "pPr");
    const level = Math.min(8, Math.max(0, Number(attr(properties, "lvl")) || 0));
    const props = mergeParagraph(listStyle[level], parseParagraphProps(properties, context));
    const base: RunStyle = { size: 18, ...props.run };
    const content = runs(paragraph, base, context);
    const firstRun = content.find((item): item is TextRun => !("type" in item));
    const numbered = props.bullet === "auto";
    const number = numbered ? (counters[level] ?? (props.bulletStart ?? 1) - 1) + 1 : undefined;
    if (firstRun) {
      counters.length = level + 1;
      counters[level] = number;
    }
    return defined({
      level,
      align: props.align ?? "left",
      marginLeft: props.marginLeft ?? 0,
      indent: props.indent ?? 0,
      lineSpacing: props.lineSpacing,
      spaceBefore: props.spaceBefore,
      spaceAfter: props.spaceAfter,
      bullet: firstRun ? bulletFor(props, number ?? 1, firstRun.style) : undefined,
      runs: content,
      endStyle: { ...base, ...parseRunProps(first(paragraph, "endParaRPr"), context) },
    }) as TextParagraph;
  });
  const vertical = bodyProps.vertical === "horizontal" ? undefined : bodyProps.vertical;
  return defined({
    anchor: bodyProps.anchor ?? "top",
    insets: {
      left: bodyProps.left ?? defaultInsets.left,
      top: bodyProps.top ?? defaultInsets.top,
      right: bodyProps.right ?? defaultInsets.right,
      bottom: bodyProps.bottom ?? defaultInsets.bottom,
    },
    wrap: bodyProps.wrap ?? true,
    vertical,
    fontScale: bodyProps.fontScale ?? 1,
    lineReduction: bodyProps.lineReduction ?? 0,
    paragraphs,
  }) as TextBody;
};
