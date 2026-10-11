import { attr, children, descendants, first } from "../../xml";
import { openOpc, type OpcPackage, type Relationship } from "../../opc/package";
import type { Block, DocxDocument, DocxSection, SectionProperties } from "../model";
import { assignListMarkers, parseNumbering } from "../numbering";
import { parseStyles } from "../styles";
import { parseThemeFonts } from "../theme";
import { twips } from "../units";
import { applyContextualSpacing, blockNodes, splitAtPageBreaks } from "./blocks";
import { parseComments } from "./comments";
import { createParseState, type ParseConfig, type ParseContext } from "./context";
import { relationshipId } from "./inline";
import { parseParagraph } from "./paragraph";
import { parseTable } from "./table";

function sectionProperties(node: Element | undefined): SectionProperties {
  const pageSize = first(node, "pgSz");
  const pageMargins = first(node, "pgMar");
  const grid = first(node, "docGrid");
  const orientation = attr(pageSize, "orient") === "landscape" ? "landscape" : "portrait";
  const rawWidth = twips(attr(pageSize, "w"), 595.3);
  const rawHeight = twips(attr(pageSize, "h"), 841.9);
  const width = orientation === "landscape" ? Math.max(rawWidth, rawHeight) : rawWidth;
  const height = orientation === "landscape" ? Math.min(rawWidth, rawHeight) : rawHeight;
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
    titlePage: !!first(node, "titlePg"),
    breakType: attr(first(node, "type"), "val") ?? "nextPage",
    linePitch:
      grid && ["lines", "linesAndChars"].includes(attr(grid, "type") ?? "")
        ? twips(attr(grid, "linePitch"), 0) || undefined
        : undefined,
  };
}

interface ParsedSection {
  properties: SectionProperties;
  references: Readonly<Record<string, string>>;
  blocks: readonly Block[];
}

function sectionReferences(node: Element): Readonly<Record<string, string>> {
  return Object.fromEntries(
    [...descendants(node, "headerReference"), ...descendants(node, "footerReference")].flatMap(
      (reference) => {
        const id = relationshipId(reference);
        const type = attr(reference, "type");
        return id && type ? [[id, type] as const] : [];
      },
    ),
  );
}

function parseBlocks(nodes: readonly Element[], context: ParseContext): Block[] {
  const blocks: Block[] = [];
  for (const node of blockNodes([...nodes])) {
    if (node.localName === "p") blocks.push(...splitAtPageBreaks(parseParagraph(node, context)));
    else if (node.localName === "tbl") blocks.push(parseTable(node, context));
  }
  return applyContextualSpacing(blocks);
}

function parseSections(nodes: readonly Element[], context: ParseContext): ParsedSection[] {
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
}

async function sectionDecorations(
  opc: OpcPackage,
  source: string,
  rels: Map<string, Relationship>,
  refs: Readonly<Record<string, string>>,
  base: Omit<ParseConfig, "source" | "rels">,
): Promise<Partial<SectionProperties>> {
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
}

export async function parseDocxModel(input: ArrayBuffer | Uint8Array): Promise<DocxDocument> {
  const opc = openOpc(input);
  const packageRels = await opc.relationships("");
  const officeDocument = [...packageRels.values()].find(
    (relation) => !relation.external && relation.type.endsWith("/officeDocument"),
  );
  const documentPath = officeDocument
    ? opc.resolve("", officeDocument.target)
    : "word/document.xml";
  const doc = await opc.xml(documentPath);
  const rels = await opc.relationships(documentPath);
  const relatedPart = (suffix: string): string | undefined => {
    const relation = [...rels.values()].find(
      (item) => !item.external && item.type.endsWith(`/${suffix}`),
    );
    return relation ? opc.resolve(documentPath, relation.target) : undefined;
  };
  const themePath = relatedPart("theme");
  const stylesPath = relatedPart("styles");
  const numberingPath = relatedPart("numbering");
  const settingsPath = relatedPart("settings") ?? "word/settings.xml";
  const theme = parseThemeFonts(
    themePath && opc.zip.has(themePath) ? await opc.xml(themePath) : undefined,
  );
  const stylesDoc = stylesPath && opc.zip.has(stylesPath) ? await opc.xml(stylesPath) : undefined;
  const numberingDoc =
    numberingPath && opc.zip.has(numberingPath) ? await opc.xml(numberingPath) : undefined;
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
  const imageParts: DocxDocument["imageParts"] = Object.fromEntries(
    await Promise.all(
      imagePaths.map(async (path) => {
        const [bytes, mime] = await Promise.all([opc.bytes(path), opc.mime(path)]);
        return [path, { bytes, mime }] as const;
      }),
    ),
  );
  const settings = opc.zip.has(settingsPath) ? await opc.xml(settingsPath) : undefined;
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
}

function collectImagePaths(
  sections: readonly DocxSection[],
  comments: readonly DocxDocument["comments"][number][],
): string[] {
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
}
