import type { Paragraph } from "../docx/model";
import type { RenderContext } from "./context";
import { renderInline } from "./inline";
import { validColor } from "./sanitize";
import { applyTextStyle } from "./text-style";

export function renderParagraph(block: Paragraph, context: RenderContext): HTMLElement {
  const paragraph = document.createElement("p");
  paragraph.className = "orim-paragraph";
  const titles: string[] = [];
  const props = block.properties;
  if (props.styleId) paragraph.classList.add(`orim-style-${props.styleId.replace(/[^\w-]/g, "")}`);
  if (props.align) {
    const align = props.align;
    paragraph.style.textAlign = ["both", "distribute"].includes(align)
      ? "justify"
      : align === "start"
        ? "left"
        : align === "end"
          ? "right"
          : align;
  }
  if (props.before !== undefined) paragraph.style.marginTop = `${props.before}pt`;
  if (props.after !== undefined) paragraph.style.marginBottom = `${props.after}pt`;
  if (props.lineRule === "auto" || props.lineRule === undefined)
    paragraph.style.lineHeight = `calc(var(--orim-pitch, 1.15) * ${props.line ?? 1})`;
  else if (props.line !== undefined) paragraph.style.lineHeight = `${props.line}pt`;
  if (props.left) paragraph.style.marginLeft = `${props.left}pt`;
  if (props.right) paragraph.style.marginRight = `${props.right}pt`;
  if (props.firstLine) paragraph.style.textIndent = `${props.firstLine}pt`;
  if (props.hanging) paragraph.style.textIndent = `${-props.hanging}pt`;
  const shading = validColor(props.shading);
  if (shading) paragraph.style.backgroundColor = shading;
  for (const side of ["top", "right", "bottom", "left"] as const) {
    const border = props.borders?.[side];
    if (border) paragraph.style.setProperty(`border-${side}`, border);
    const padding = props.borderSpacing?.[side];
    if (padding !== undefined) paragraph.style.setProperty(`padding-${side}`, `${padding}pt`);
  }
  if (props.snapToGrid === false) paragraph.style.setProperty("--orim-pitch", "initial");
  if (props.change) {
    paragraph.classList.add(
      "orim-paragraph--change",
      `orim-paragraph--change-${props.change.kind}`,
    );
    titles.push(
      `Paragraph mark ${props.change.kind === "del" ? "deleted" : "inserted"} by ${props.change.author}${props.change.date ? ` on ${props.change.date}` : ""}`,
    );
  }
  if (props.formattingChange) {
    paragraph.classList.add("orim-format-change");
    titles.push(
      `Formatted: ${props.formattingChange.author}${props.formattingChange.date ? ` ${props.formattingChange.date}` : ""}`,
    );
  }
  if (titles.length) paragraph.title = titles.join("\n");
  if (props.pageBreakBefore) paragraph.style.breakBefore = "page";
  if (block.list) {
    paragraph.classList.add("orim-list");
    paragraph.style.marginLeft = `${Math.max(0, block.list.left ?? 18)}pt`;
    paragraph.style.textIndent = `${-(block.list.hanging ?? 18)}pt`;
    const marker = document.createElement("span");
    marker.className = "orim-list-marker";
    marker.textContent =
      block.list.marker +
      (block.list.suffix === "space" ? " " : block.list.suffix === "tab" ? "\t" : "");
    marker.style.minWidth = `${block.list.hanging ?? 18}pt`;
    applyTextStyle(marker, block.markStyle);
    paragraph.append(marker);
  }
  for (const inline of block.runs) paragraph.append(renderInline(inline, context));
  const hasVisibleInline = block.runs.some(
    (inline) =>
      (inline.type === "text" && inline.text.length > 0 && !inline.style?.hidden) ||
      inline.type === "image" ||
      inline.type === "field" ||
      (inline.type === "comment-reference" && context.comments),
  );
  if (!hasVisibleInline) {
    const spacer = document.createElement("span");
    spacer.textContent = "\u200b";
    applyTextStyle(spacer, block.markStyle);
    paragraph.append(spacer);
  }
  if (props.change) {
    const mark = document.createElement("span");
    mark.className = "orim-paragraph-mark";
    mark.textContent = " ¶";
    paragraph.append(mark);
  }
  return paragraph;
}
