import { attr, children, descendants, first } from "../xml";
import {
  findRelationship,
  loadOpc,
  readMedia,
  type OpcPackage,
  type PackageInput,
  type Relationship,
} from "../../opc/package";
import { themeFonts } from "../../drawingml/theme";
import type { Block, DocxDocument, DocxSection, SectionProperties } from "../model";
import { assignListMarkers, parseNumbering } from "../numbering";
import { parseStyles } from "../styles";
import { twips } from "../units";
import { applyContextualSpacing, blockNodes, splitAtPageBreaks } from "./blocks";
import { parseComments } from "./comments";
import { createParseState, type ParseConfig, type ParseContext } from "./context";
import { relationshipId } from "./inline";
import { parseParagraph } from "./paragraph";
import { sectionProperties } from "./section";
import { parseTable } from "./table";

interface ParsedSection {
  properties: SectionProperties;
  references: Readonly<Record<string, string>>;
  blocks: readonly Block[];
}

const sectionReferences = (node: Element): Readonly<Record<string, string>> => {
  return Object.fromEntries(
    [...descendants(node, "headerReference"), ...descendants(node, "footerReference")].flatMap(
      (reference) => {
        const id = relationshipId(reference);
        const type = attr(reference, "type");
        return id && type ? [[id, type] as const] : [];
      },
    ),
  );
};

const parseBlocks = (nodes: readonly Element[], context: ParseContext): Block[] => {
  const blocks: Block[] = [];
  for (const node of blockNodes([...nodes])) {
    if (node.localName === "p") blocks.push(...splitAtPageBreaks(parseParagraph(node, context)));
    else if (node.localName === "tbl") blocks.push(parseTable(node, context));
  }
  return applyContextualSpacing(blocks);
};

const parseSections = (nodes: readonly Element[], context: ParseContext): ParsedSection[] => {
  const sections: ParsedSection[] = [];
  let blocks: Block[] = [];
  const pushSection = (
    properties: SectionProperties,
    references: Readonly<Record<string, string>>,
  ) => {
    sections.push({ properties, references, blocks: applyContextualSpacing(blocks) });
    blocks = [];
  };

  for (const node of nodes) {
    if (node.localName === "sectPr") {
      pushSection(sectionProperties(node), sectionReferences(node));
      continue;
    }
    if (node.localName !== "p" && node.localName !== "tbl") continue;

    if (node.localName === "p") blocks.push(...splitAtPageBreaks(parseParagraph(node, context)));
    else blocks.push(parseTable(node, context));

    const paragraphSection =
      node.localName === "p" ? first(first(node, "pPr"), "sectPr") : undefined;
    if (paragraphSection)
      pushSection(sectionProperties(paragraphSection), sectionReferences(paragraphSection));
  }

  if (blocks.length) {
    const previous = sections.at(-1);
    pushSection(previous?.properties ?? sectionProperties(undefined), previous?.references ?? {});
  }
  return sections;
};

const sectionDecorations = async (
  opc: OpcPackage,
  source: string,
  rels: Map<string, Relationship>,
  refs: Readonly<Record<string, string>>,
  base: Omit<ParseConfig, "source" | "rels">,
): Promise<Partial<SectionProperties>> => {
  const parsed = await Promise.all(
    [...rels.values()].map(async (rel) => {
      const type = rel.type.slice(rel.type.lastIndexOf("/") + 1);
      const kind = refs[rel.id];
      if ((type !== "header" && type !== "footer") || !kind) return undefined;
      const part = opc.resolve(source, rel.target);
      if (!opc.zip.has(part)) return undefined;
      const doc = await opc.xml(part);
      const root = doc.documentElement;
      if (!root) return undefined;
      const partContext: ParseContext = {
        config: { ...base, source: part, rels: await opc.relationships(part) },
        state: createParseState(),
      };
      return { type, kind, blocks: parseBlocks(children(root), partContext) } as const;
    }),
  );
  const decorations = (type: "header" | "footer") =>
    Object.fromEntries(
      parsed.flatMap((item) => (item?.type === type ? [[item.kind, item.blocks] as const] : [])),
    );
  const headers = decorations("header");
  const footers = decorations("footer");
  return {
    header: headers.default,
    firstHeader: headers.first,
    evenHeader: headers.even,
    footer: footers.default,
    firstFooter: footers.first,
    evenFooter: footers.even,
  };
};

export const parseDocxModel = async (input: PackageInput): Promise<DocxDocument> => {
  const opc = await loadOpc(input);
  const documentPath = (await opc.part("", "officeDocument")) ?? "word/document.xml";
  const doc = await opc.xml(documentPath);
  if (doc.documentElement.localName !== "document")
    throw new Error("Not a DOCX package: document part is not a document.");
  const rels = await opc.relationships(documentPath);
  const relatedPart = (type: string): string | undefined => {
    const relation = findRelationship(rels, type);
    return relation && opc.resolve(documentPath, relation.target);
  };
  const themePath = relatedPart("theme");
  const stylesPath = relatedPart("styles");
  const numberingPath = relatedPart("numbering");
  const settingsPath = relatedPart("settings") ?? "word/settings.xml";
  const theme = themeFonts(await opc.optionalXml(themePath));
  const stylesDoc = await opc.optionalXml(stylesPath);
  const numberingDoc = await opc.optionalXml(numberingPath);
  const styles = parseStyles(stylesDoc, theme);
  const numbering = parseNumbering(numberingDoc);
  const context: ParseContext = {
    config: { opc, source: documentPath, rels, styles, theme, numbering },
    state: createParseState(),
  };
  const body = first(doc.documentElement, "body");
  const sections = parseSections(blockNodes(children(body ?? doc)), context);
  const renderedSections: DocxSection[] = [];
  for (const section of sections) {
    const previous = renderedSections.at(-1)?.properties;
    const decorations = await sectionDecorations(opc, documentPath, rels, section.references, {
      opc,
      styles,
      theme,
      numbering,
    });
    const properties: SectionProperties = {
      ...section.properties,
      ...decorations,
      header: decorations.header ?? previous?.header,
      firstHeader: decorations.firstHeader ?? previous?.firstHeader,
      evenHeader: decorations.evenHeader ?? previous?.evenHeader,
      footer: decorations.footer ?? previous?.footer,
      firstFooter: decorations.firstFooter ?? previous?.firstFooter,
      evenFooter: decorations.evenFooter ?? previous?.evenFooter,
    };
    renderedSections.push({ properties, blocks: section.blocks });
  }

  const bodyBlocks = renderedSections.flatMap((section) => section.blocks);
  const numberedBodyBlocks = assignListMarkers(bodyBlocks, numbering);
  let bodyBlockIndex = 0;
  const numberedSections = renderedSections.map((section) => {
    const blocks = numberedBodyBlocks.slice(bodyBlockIndex, bodyBlockIndex + section.blocks.length);
    bodyBlockIndex += section.blocks.length;
    const numberStory = (story: readonly Block[] | undefined) =>
      story ? assignListMarkers(story, numbering) : undefined;
    return {
      ...section,
      properties: {
        ...section.properties,
        header: numberStory(section.properties.header),
        firstHeader: numberStory(section.properties.firstHeader),
        evenHeader: numberStory(section.properties.evenHeader),
        footer: numberStory(section.properties.footer),
        firstFooter: numberStory(section.properties.firstFooter),
        evenFooter: numberStory(section.properties.evenFooter),
      },
      blocks,
    };
  });
  const commentList = (await parseComments(opc, context, parseBlocks)).map((comment) => ({
    ...comment,
    paragraphs: assignListMarkers(comment.paragraphs, numbering),
  }));
  const imagePaths = collectImagePaths(numberedSections, commentList);
  const imageParts = await readMedia(opc, imagePaths);
  const settings = await opc.optionalXml(settingsPath);
  const defaultTabStop = settings
    ? twips(attr(descendants(settings, "defaultTabStop")[0], "val"), 36)
    : undefined;
  return {
    sections: numberedSections,
    comments: commentList,
    evenAndOddHeaders: !!(settings && descendants(settings, "evenAndOddHeaders").length),
    defaultTabStop,
    imageParts,
  };
};

const collectImagePaths = (
  sections: readonly DocxSection[],
  comments: readonly DocxDocument["comments"][number][],
): string[] => {
  const paths = new Set<string>();
  const visit = (blocks: readonly Block[]) => {
    for (const block of blocks) {
      if (block.type === "paragraph") {
        for (const run of block.runs) if (run.type === "image") paths.add(run.path);
      } else {
        for (const row of block.rows) for (const cell of row.cells) visit(cell.blocks);
      }
    }
  };
  for (const section of sections) {
    visit(section.blocks);
    visit(section.properties.header ?? []);
    visit(section.properties.firstHeader ?? []);
    visit(section.properties.evenHeader ?? []);
    visit(section.properties.footer ?? []);
    visit(section.properties.firstFooter ?? []);
    visit(section.properties.evenFooter ?? []);
  }
  for (const comment of comments) visit(comment.paragraphs);
  return [...paths];
};
