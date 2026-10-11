import { attr, children, descendants, first } from "../../xml";
import { openOpc, type OpcPackage, type Relationship } from "../../opc/package";
import type { Block, DocxDocument, SectionProperties } from "../model";
import { parseNumbering } from "../numbering";
import { parseStyles } from "../styles";
import { twips } from "../units";
import { applyContextualSpacing, blockNodes, splitAtPageBreaks } from "./blocks";
import { parseComments } from "./comments";
import type { ParseContext } from "./context";
import { createFieldStack } from "./fields";
import { relationshipId } from "./inline";
import { parseParagraph } from "./paragraph";
import { parseTable } from "./table";

function sectionProperties(node: Element | undefined): SectionProperties {
  const pageSize = first(node ?? document, "pgSz");
  const pageMargins = first(node ?? document, "pgMar");
  const grid = first(node ?? document, "docGrid");
  const orientation = attr(pageSize, "orient") === "landscape" ? "landscape" : "portrait";
  let width = twips(attr(pageSize, "w"), 595.3);
  let height = twips(attr(pageSize, "h"), 841.9);
  if (orientation === "landscape" && width < height) [width, height] = [height, width];
  return {
    width,
    height,
    orientation,
    margins: {
      top: twips(attr(pageMargins, "top"), 1440 / 20),
      right: twips(attr(pageMargins, "right"), 1440 / 20),
      bottom: twips(attr(pageMargins, "bottom"), 1440 / 20),
      left: twips(attr(pageMargins, "left"), 1440 / 20),
      header: twips(attr(pageMargins, "header"), 720 / 20),
      footer: twips(attr(pageMargins, "footer"), 720 / 20),
    },
    titlePage: !!first(node ?? document, "titlePg"),
    breakType: attr(first(node ?? document, "type"), "val") ?? "nextPage",
    linePitch: grid ? twips(attr(grid, "linePitch"), 0) || undefined : undefined,
  };
}

function parseBlocks(nodes: Element[], context: ParseContext, comments: Set<string>): Block[] {
  const out: Block[] = [];
  for (const node of blockNodes(nodes)) {
    if (node.localName === "p")
      out.push(...splitAtPageBreaks(parseParagraph(node, context, comments)));
    else if (node.localName === "tbl") out.push(parseTable(node, context, comments));
  }
  return applyContextualSpacing(out);
}

async function sectionDecorations(
  opc: OpcPackage,
  properties: SectionProperties,
  source: string,
  rels: Map<string, Relationship>,
  base: Omit<ParseContext, "source" | "rels" | "commentsOpen">,
): Promise<void> {
  const headers: Record<string, Block[]> = {};
  const footers: Record<string, Block[]> = {};
  const refs = (properties as SectionProperties & { _refs?: Record<string, string> })._refs ?? {};
  for (const rel of rels.values()) {
    const type = rel.type.slice(rel.type.lastIndexOf("/") + 1);
    if (type !== "header" && type !== "footer") continue;
    if (!refs[rel.id]) continue;
    const part = opc.resolve(source, rel.target);
    if (!opc.zip.has(part)) continue;
    const doc = await opc.xml(part);
    const root = doc.documentElement;
    if (!root) continue;
    const partContext: ParseContext = {
      ...base,
      source: part,
      rels: await opc.relationships(part),
      commentsOpen: new Set(),
      fields: createFieldStack(),
    };
    const blocks = parseBlocks(children(root), partContext, new Set());
    const sectionRef = rel.id;
    const kind = refs[sectionRef ?? ""] ?? "default";
    (type === "header" ? headers : footers)[kind] = blocks;
  }
  properties.header = headers.default;
  properties.firstHeader = headers.first;
  properties.evenHeader = headers.even;
  properties.footer = footers.default;
  properties.firstFooter = footers.first;
  properties.evenFooter = footers.even;
}

export async function parseDocxModel(input: ArrayBuffer | Uint8Array): Promise<DocxDocument> {
  const opc = openOpc(input);
  const documentPath = "word/document.xml";
  const doc = await opc.xml(documentPath);
  const rels = await opc.relationships(documentPath);
  const theme = await opc.themeFonts();
  const stylesDoc = opc.zip.has("word/styles.xml") ? await opc.xml("word/styles.xml") : undefined;
  const numberingDoc = opc.zip.has("word/numbering.xml")
    ? await opc.xml("word/numbering.xml")
    : undefined;
  const styles = parseStyles(stylesDoc, theme);
  const numbering = parseNumbering(numberingDoc);
  const commentsOpen = new Set<string>();
  const images = new Set<string>();
  const context: ParseContext = {
    opc,
    source: documentPath,
    rels,
    styles,
    theme,
    numbering,
    commentsOpen,
    images,
    fields: createFieldStack(),
  };
  const body = first(doc.documentElement, "body");
  const sections: DocxDocument["sections"] = [];
  let blocks: Block[] = [];
  let secPr: Element | undefined;
  for (const node of blockNodes(children(body ?? doc))) {
    if (node.localName === "sectPr") {
      secPr = node;
      const props = sectionProperties(secPr);
      const refs: Record<string, string> = {};
      for (const ref of [
        ...descendants(secPr, "headerReference"),
        ...descendants(secPr, "footerReference"),
      ]) {
        const id = relationshipId(ref);
        const type = attr(ref, "type");
        if (id && type) refs[id] = type;
      }
      (props as SectionProperties & { _refs?: Record<string, string> })._refs = refs;
      sections.push({ properties: props, blocks: applyContextualSpacing(blocks) });
      blocks = [];
      secPr = undefined;
    } else if (node.localName === "p" || node.localName === "tbl") {
      if (node.localName === "p")
        blocks.push(...splitAtPageBreaks(parseParagraph(node, context, commentsOpen)));
      else blocks.push(parseTable(node, context, commentsOpen));
      const pPr = node.localName === "p" ? first(node, "pPr") : undefined;
      const paragraphSection = first(pPr ?? node, "sectPr");
      if (paragraphSection) {
        const props = sectionProperties(paragraphSection);
        const refs: Record<string, string> = {};
        for (const ref of [
          ...descendants(paragraphSection, "headerReference"),
          ...descendants(paragraphSection, "footerReference"),
        ]) {
          const id = relationshipId(ref);
          const type = attr(ref, "type");
          if (id && type) refs[id] = type;
        }
        (props as SectionProperties & { _refs?: Record<string, string> })._refs = refs;
        sections.push({ properties: props, blocks: applyContextualSpacing(blocks) });
        blocks = [];
      }
    }
  }
  if (blocks.length || !sections.length)
    sections.push({ properties: sectionProperties(secPr), blocks: applyContextualSpacing(blocks) });

  for (let index = 0; index < sections.length; index++) {
    const section = sections[index];
    const props = section.properties as SectionProperties & { _refs?: Record<string, string> };
    await sectionDecorations(opc, props, documentPath, rels, {
      opc,
      styles,
      theme,
      numbering,
      images,
      fields: createFieldStack(),
    });
    const previous = sections[index - 1]?.properties;
    if (previous) {
      props.header ??= previous.header;
      props.firstHeader ??= previous.firstHeader;
      props.evenHeader ??= previous.evenHeader;
      props.footer ??= previous.footer;
      props.firstFooter ??= previous.firstFooter;
      props.evenFooter ??= previous.evenFooter;
    }
  }

  const commentList = await parseComments(opc, context, parseBlocks);
  const imageParts: DocxDocument["imageParts"] = {};
  for (const path of images)
    imageParts[path] = { bytes: await opc.bytes(path), mime: await opc.mime(path) };
  const settings = opc.zip.has("word/settings.xml")
    ? await opc.xml("word/settings.xml")
    : undefined;
  const defaultTabStop = settings
    ? twips(attr(descendants(settings, "defaultTabStop")[0], "val"), 36)
    : undefined;
  return {
    sections,
    comments: commentList,
    evenAndOddHeaders: !!(settings && descendants(settings, "evenAndOddHeaders").length),
    defaultTabStop,
    imageParts,
  };
}
