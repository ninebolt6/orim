import type { Block, Inline, Paragraph } from "../model";
import { children, elements, first } from "../../xml";

export const splitAtPageBreaks = (paragraph: Paragraph): Paragraph[] => {
  if (!paragraph.runs.some((run) => run.type === "break" && run.page)) return [paragraph];
  const segments: Array<{ runs: Inline[]; pageBreakAfter: boolean }> = [];
  let current: Inline[] = [];
  let leadingBreak = false;
  for (const run of paragraph.runs) {
    if (run.type !== "break" || !run.page) {
      current.push(run);
    } else if (current.length) {
      segments.push({ runs: current, pageBreakAfter: true });
      current = [];
    } else if (segments.length) {
      segments[segments.length - 1] = { ...segments[segments.length - 1], pageBreakAfter: true };
    } else {
      leadingBreak = true;
    }
  }
  if (current.length) segments.push({ runs: current, pageBreakAfter: false });
  if (!segments.length)
    return [
      {
        ...paragraph,
        properties: { ...paragraph.properties, pageBreakAfter: true },
        runs: [],
      },
    ];
  return segments.map(({ runs: segmentRuns, pageBreakAfter }, index) => ({
    ...paragraph,
    properties: {
      ...paragraph.properties,
      before: index ? 0 : paragraph.properties.before,
      pageBreakBefore:
        index === 0 && leadingBreak ? true : index ? false : paragraph.properties.pageBreakBefore,
      pageBreakAfter,
    },
    runs: segmentRuns,
    list: index === 0 ? paragraph.list : undefined,
  }));
};

export const applyContextualSpacing = (blocks: readonly Block[]): Block[] => {
  return blocks.map((block, index) => {
    if (block.type !== "paragraph") return block;
    const previous = blocks[index - 1];
    const next = blocks[index + 1];
    const sameAsPrevious =
      previous?.type === "paragraph" && previous.properties.styleId === block.properties.styleId;
    const sameAsNext =
      next?.type === "paragraph" && next.properties.styleId === block.properties.styleId;
    return {
      ...block,
      properties: {
        ...block.properties,
        ...(sameAsPrevious && block.properties.contextualSpacing ? { before: 0 } : {}),
        ...(sameAsNext && block.properties.contextualSpacing ? { after: 0 } : {}),
      },
    };
  });
};

export const blockNodes = (nodes: Element[]): Element[] => {
  const out: Element[] = [];
  for (const node of nodes) {
    if (["p", "tbl", "sectPr"].includes(node.localName)) out.push(node);
    else if (node.localName === "AlternateContent") {
      const fallback = elements(node, "Fallback")[0];
      if (fallback) out.push(...blockNodes(children(fallback)));
    } else if (node.localName === "sdt") {
      const content = first(node, "sdtContent");
      if (content) out.push(...blockNodes(children(content)));
    } else if (["sdtContent", "customXml", "smartTag", "Fallback"].includes(node.localName)) {
      out.push(...blockNodes(children(node)));
    }
  }
  return out;
};
