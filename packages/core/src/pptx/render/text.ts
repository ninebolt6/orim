import { safeHref, validColor } from "../../common/sanitize";
import type { RunStyle, Spacing, TextBody, TextParagraph } from "../model";
import { fontFamily, pt } from "./style";

const applyRunStyle = (element: HTMLElement, style: RunStyle, scale: number): void => {
  const size = (style.size ?? 18) * scale * (style.baseline ? 2 / 3 : 1);
  element.style.fontSize = pt(size);
  if (style.bold) element.style.fontWeight = "bold";
  if (style.italic) element.style.fontStyle = "italic";
  const decorations = [style.underline && "underline", style.strike && "line-through"].filter(
    Boolean,
  );
  if (decorations.length) element.style.textDecorationLine = decorations.join(" ");
  const color = validColor(style.color);
  if (color) element.style.color = color;
  const highlight = validColor(style.highlight);
  if (highlight) element.style.backgroundColor = highlight;
  if (style.font || style.eastAsianFont)
    element.style.fontFamily = fontFamily(style.font, style.eastAsianFont);
  if (style.baseline)
    element.style.verticalAlign = `${Number((style.baseline * 1.5).toFixed(3))}em`;
  if (style.caps === "all") element.style.textTransform = "uppercase";
  if (style.caps === "small") element.style.fontVariant = "small-caps";
  if (style.spacing) element.style.letterSpacing = pt(style.spacing);
};

const spacing = (value: Spacing | undefined, reduction = 0): string | undefined =>
  value === undefined
    ? undefined
    : "points" in value
      ? pt(value.points)
      : `${Number((value.percent * (1 - reduction)).toFixed(4))}em`;

const renderParagraph = (paragraph: TextParagraph, body: TextBody): HTMLElement => {
  const element = document.createElement("p");
  element.className = "orim-pptx-paragraph";
  const firstRun = paragraph.runs.find((run) => !("type" in run));
  const baseSize = (firstRun?.style.size ?? paragraph.endStyle.size ?? 18) * body.fontScale;
  element.style.fontSize = pt(baseSize);
  element.style.textAlign = paragraph.align;
  if (paragraph.marginLeft) element.style.paddingLeft = pt(paragraph.marginLeft);
  if (paragraph.indent) element.style.textIndent = pt(paragraph.indent);
  const line = paragraph.lineSpacing;
  if (line)
    element.style.lineHeight =
      "points" in line
        ? pt(line.points * (1 - body.lineReduction))
        : String(Number((1.2 * line.percent * (1 - body.lineReduction)).toFixed(4)));
  else if (body.lineReduction)
    element.style.lineHeight = String(Number((1.2 * (1 - body.lineReduction)).toFixed(4)));
  const before = spacing(paragraph.spaceBefore);
  const after = spacing(paragraph.spaceAfter);
  if (before) element.style.marginTop = before;
  if (after) element.style.marginBottom = after;
  if (paragraph.bullet && firstRun) {
    const bullet = document.createElement("span");
    bullet.className = "orim-pptx-bullet";
    bullet.append(document.createTextNode(paragraph.bullet.text));
    bullet.style.fontSize = pt(
      paragraph.bullet.size ??
        (firstRun.style.size ?? 18) * body.fontScale * (paragraph.bullet.scale ?? 1),
    );
    const color = validColor(paragraph.bullet.color);
    if (color) bullet.style.color = color;
    if (paragraph.bullet.font) bullet.style.fontFamily = fontFamily(paragraph.bullet.font);
    if (paragraph.indent < 0) bullet.style.minWidth = pt(-paragraph.indent);
    else bullet.style.marginRight = "0.5em";
    element.append(bullet);
  }
  for (const run of paragraph.runs) {
    if ("type" in run) {
      element.append(document.createElement("br"));
      continue;
    }
    const span = document.createElement("span");
    applyRunStyle(span, run.style, body.fontScale);
    span.append(document.createTextNode(run.text));
    const href = safeHref(run.href);
    if (href) {
      const link = document.createElement("a");
      link.href = href;
      if (!href.startsWith("#")) link.rel = "noopener noreferrer";
      link.append(span);
      element.append(link);
    } else element.append(span);
  }
  if (!firstRun) element.append(document.createElement("br"));
  return element;
};

export const renderParagraphs = (body: TextBody, parent: HTMLElement): void => {
  for (const paragraph of body.paragraphs) parent.append(renderParagraph(paragraph, body));
};

export const renderTextBody = (body: TextBody): HTMLElement => {
  const element = document.createElement("div");
  element.className = `orim-pptx-text orim-pptx-text--${body.anchor}`;
  if (!body.wrap) element.classList.add("orim-pptx-text--nowrap");
  if (body.vertical) element.classList.add(`orim-pptx-text--${body.vertical}`);
  element.style.padding = [body.insets.top, body.insets.right, body.insets.bottom, body.insets.left]
    .map(pt)
    .join(" ");
  renderParagraphs(body, element);
  return element;
};
