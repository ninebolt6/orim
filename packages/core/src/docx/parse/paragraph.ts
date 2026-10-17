import { attr, first } from "../xml";
import type { Paragraph, ParagraphProperties } from "../model";
import { parseRunProperties, paragraphProperties } from "../styles";
import type { ParseContext } from "./context";
import { changeFrom, collectInline, revisionInfoFrom } from "./inline";

export const parseParagraph = (node: Element, context: ParseContext): Paragraph => {
  const { styles, theme, numbering } = context.config;
  const pPr = first(node, "pPr");
  const pStyle = attr(first(pPr, "pStyle"), "val");
  const directProperties = paragraphProperties(pPr);
  const markChange = first(pPr, "ins") ?? first(pPr, "del");
  const formattingChange = first(pPr, "pPrChange");
  const properties: ParagraphProperties = {
    ...styles.paragraph(pStyle, directProperties),
    styleId: pStyle,
    ...(markChange ? { change: changeFrom(markChange, first(pPr, "del") ? "del" : "ins") } : {}),
    ...(formattingChange ? { formattingChange: revisionInfoFrom(formattingChange) } : {}),
  };
  const numPr = first(pPr, "numPr");
  const inheritedNumbering = styles.numbering(pStyle);
  const numId = attr(first(numPr, "numId"), "val") ?? inheritedNumbering.numId;
  const level = Number(attr(first(numPr, "ilvl"), "val") ?? inheritedNumbering.level ?? 0);
  const runs = collectInline(node, {
    ...context,
    config: { ...context.config, paragraphStyle: pStyle },
  });
  const marker = numId && numId !== "0" ? numbering.level(numId, level) : undefined;
  const inheritedProperties = styles.paragraph(pStyle);
  return {
    type: "paragraph",
    properties,
    runs,
    markStyle: styles.run(pStyle, undefined, parseRunProperties(first(pPr, "rPr"), theme)),
    list: marker
      ? {
          numId: numId!,
          level,
          left: directProperties.left ?? marker.left ?? inheritedProperties.left,
          hanging: directProperties.hanging ?? marker.hanging ?? inheritedProperties.hanging,
          suffix: marker.suffix,
        }
      : undefined,
  };
};
