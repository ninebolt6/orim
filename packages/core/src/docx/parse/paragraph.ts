import { attr, first } from "../../xml";
import type { Paragraph, ParagraphProperties } from "../model";
import { parseRunProperties, paragraphProperties } from "../styles";
import type { ParseContext } from "./context";
import { changeFrom, collectInline } from "./inline";

export function parseParagraph(
  node: Element,
  context: ParseContext,
  commentState: Set<string>,
): Paragraph {
  const pPr = first(node, "pPr");
  const pStyle = attr(first(pPr ?? node, "pStyle"), "val");
  const directProperties = paragraphProperties(pPr) as Partial<ParagraphProperties>;
  const properties = context.styles.paragraph(pStyle, directProperties);
  properties.styleId = pStyle;
  const numPr = first(pPr ?? node, "numPr");
  const inheritedNumbering = context.styles.numbering(pStyle);
  const numId = attr(first(numPr ?? node, "numId"), "val") ?? inheritedNumbering.numId;
  const level = Number(attr(first(numPr ?? node, "ilvl"), "val") ?? inheritedNumbering.level ?? 0);
  const runs = collectInline(node, { ...context, paragraphStyle: pStyle }, commentState);
  const marker = numId && numId !== "0" ? context.numbering.marker(numId, level) : undefined;
  const inheritedProperties = context.styles.paragraph(pStyle);
  const markChange = first(pPr ?? node, "ins") ?? first(pPr ?? node, "del");
  if (markChange)
    properties.change = changeFrom(markChange, first(pPr ?? node, "del") ? "del" : "ins");
  const formattingChange = first(pPr ?? node, "pPrChange");
  if (formattingChange) properties.formattingChange = changeFrom(formattingChange, "ins");
  return {
    type: "paragraph",
    properties,
    runs,
    markStyle: context.styles.run(
      pStyle,
      undefined,
      parseRunProperties(first(pPr ?? node, "rPr"), context.theme),
    ),
    list: marker
      ? {
          numId: numId!,
          level,
          marker: marker.text,
          left: directProperties.left ?? marker.left ?? inheritedProperties.left,
          hanging: directProperties.hanging ?? marker.hanging ?? inheritedProperties.hanging,
          suffix: marker.suffix,
        }
      : undefined,
  };
}
