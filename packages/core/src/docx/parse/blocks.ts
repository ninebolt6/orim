import type { Block, Inline, Paragraph } from "../model";
import { children, elements, first } from "../../xml";

export function splitAtPageBreaks(paragraph: Paragraph): Paragraph[] {
  if (!paragraph.runs.some((run) => run.type === "break" && run.page)) return [paragraph];
  const segments: Array<{ runs: Inline[]; pageBreakAfter: boolean }> = [];
  let runs: Inline[] = [];
  let leadingBreak = false;
  for (const run of paragraph.runs) {
    if (run.type === "break" && run.page) {
      if (runs.length) {
        segments.push({ runs, pageBreakAfter: true });
        runs = [];
      } else if (!segments.length) leadingBreak = true;
      else segments.at(-1)!.pageBreakAfter = true;
    } else runs.push(run);
  }
  if (runs.length) segments.push({ runs, pageBreakAfter: false });
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
}

export function applyContextualSpacing(blocks: Block[]): Block[] {
  let previous: Paragraph | undefined;
  for (const block of blocks) {
    if (block.type !== "paragraph") {
      previous = undefined;
      continue;
    }
    if (previous && previous.properties.styleId === block.properties.styleId) {
      if (block.properties.contextualSpacing) block.properties.before = 0;
      if (previous.properties.contextualSpacing) previous.properties.after = 0;
    }
    previous = block;
  }
  return blocks;
}

export function blockNodes(nodes: Element[]): Element[] {
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
}
