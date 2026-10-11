import { attr, children, descendants, first } from "../../xml";
import { openOpc, type OpcPackage, type Relationship } from "../../opc/package";
import type { Block, DocxDocument, DocxSection, SectionProperties } from "../model";
import { parseNumbering } from "../numbering";
import { parseStyles } from "../styles";
import { twips } from "../units";
import { applyContextualSpacing, blockNodes, splitAtPageBreaks } from "./blocks";
import { parseComments } from "./comments";
import { createCommentRangeState, createImagePathState, type ParseContext } from "./context";
import { createFieldStack } from "./fields";
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
    linePitch: grid ? twips(attr(grid, "linePitch"), 0) || undefined : undefined,
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
  return applyContextualSpacing(
    blockNodes([...nodes]).flatMap((node): Block[] =>
      node.localName === "p"
        ? splitAtPageBreaks(parseParagraph(node, context))
        : node.localName === "tbl"
          ? [parseTable(node, context)]
          : [],
    ),
  );
}

function parseSections(nodes: readonly Element[], context: ParseContext): ParsedSection[] {
  const grouped = nodes.reduce(
    (state, node) => {
      if (node.localName === "sectPr") {
        return {
          sections: [
            ...state.sections,
            {
              properties: sectionProperties(node),
              references: sectionReferences(node),
              blocks: applyContextualSpacing(state.blocks),
            },
          ],
          blocks: [],
        };
      }
      if (node.localName !== "p" && node.localName !== "tbl") return state;

      const blocks = [
        ...state.blocks,
        ...(node.localName === "p"
          ? splitAtPageBreaks(parseParagraph(node, context))
          : [parseTable(node, context)]),
      ];
      const paragraphSection =
        node.localName === "p" ? first(first(node, "pPr"), "sectPr") : undefined;
      if (!paragraphSection) return { ...state, blocks };
      return {
        sections: [
          ...state.sections,
          {
            properties: sectionProperties(paragraphSection),
            references: sectionReferences(paragraphSection),
            blocks: applyContextualSpacing(blocks),
          },
        ],
        blocks: [],
      };
    },
    { sections: [] as ParsedSection[], blocks: [] as Block[] },
  );

  return grouped.blocks.length || !grouped.sections.length
    ? [
        ...grouped.sections,
        {
          properties: sectionProperties(undefined),
          references: {},
          blocks: applyContextualSpacing(grouped.blocks),
        },
      ]
    : grouped.sections;
}

async function sectionDecorations(
  opc: OpcPackage,
  source: string,
  rels: Map<string, Relationship>,
  refs: Readonly<Record<string, string>>,
  base: Omit<ParseContext, "source" | "rels" | "commentRanges">,
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
        ...base,
        source: part,
        rels: await opc.relationships(part),
        commentRanges: createCommentRangeState(),
        fields: createFieldStack(),
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
  const commentRanges = createCommentRangeState();
  const imagePartPaths = createImagePathState();
  const context: ParseContext = {
    opc,
    source: documentPath,
    rels,
    styles,
    theme,
    numbering,
    commentRanges,
    imageParts: imagePartPaths,
    fields: createFieldStack(),
  };
  const body = first(doc.documentElement, "body");
  const sections = parseSections(blockNodes(children(body ?? doc)), context);
  const renderedSections = await sections.reduce<Promise<DocxSection[]>>(
    async (promise, section) => {
      const previousSections = await promise;
      const previous = previousSections.at(-1)?.properties;
      const decorations = await sectionDecorations(opc, documentPath, rels, section.references, {
        opc,
        styles,
        theme,
        numbering,
        imageParts: imagePartPaths,
        fields: createFieldStack(),
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
      return [...previousSections, { properties, blocks: section.blocks }];
    },
    Promise.resolve([]),
  );

  const commentList = await parseComments(opc, context, parseBlocks);
  const imageParts: DocxDocument["imageParts"] = Object.fromEntries(
    await Promise.all(
      [...context.imageParts.values()].map(async (path) => {
        const [bytes, mime] = await Promise.all([opc.bytes(path), opc.mime(path)]);
        return [path, { bytes, mime }] as const;
      }),
    ),
  );
  const settings = opc.zip.has("word/settings.xml")
    ? await opc.xml("word/settings.xml")
    : undefined;
  const defaultTabStop = settings
    ? twips(attr(descendants(settings, "defaultTabStop")[0], "val"), 36)
    : undefined;
  return {
    sections: renderedSections,
    comments: commentList,
    evenAndOddHeaders: !!(settings && descendants(settings, "evenAndOddHeaders").length),
    defaultTabStop,
    imageParts,
  };
}
