import type {
  Block,
  ChangeKind,
  DocxDocument,
  Paragraph,
  SectionProperties,
  Table,
} from "../docx/model";
import type { ChangesMode } from "./document";

export const revisionVisibility = (
  kind: ChangeKind,
  mode: ChangesMode,
): "hidden" | "marked" | "plain" => {
  if (mode === "markup") return "marked";
  const accepted = kind === "ins" || kind === "moveTo";
  if (mode === "final") return accepted ? "plain" : "hidden";
  return accepted ? "hidden" : "plain";
};

const resolveTable = (table: Table, mode: ChangesMode): Table => {
  return {
    ...table,
    rows: table.rows.map((row) => ({
      ...row,
      cells: row.cells.map((cell) => ({
        ...cell,
        blocks: resolveParagraphMarks(cell.blocks, mode),
      })),
    })),
  };
};

export const resolveParagraphMarks = (blocks: readonly Block[], mode: ChangesMode): Block[] => {
  const resolved = blocks.map((block) =>
    block.type === "table" ? resolveTable(block, mode) : block,
  );
  if (mode === "markup") return resolved;

  const output: Block[] = [];
  for (let index = 0; index < resolved.length; index++) {
    const block = resolved[index];
    const visibility =
      block.type === "paragraph" && block.properties.change
        ? revisionVisibility(block.properties.change.kind, mode)
        : "plain";
    if (block.type !== "paragraph" || visibility !== "hidden") {
      output.push(block);
      continue;
    }

    let current: Paragraph = block;
    const runs = [...block.runs];
    let nextIndex = index + 1;
    while (resolved[nextIndex]?.type === "paragraph") {
      const next = resolved[nextIndex] as Paragraph;
      runs.push(...next.runs);
      current = next;
      nextIndex++;
      const nextVisibility = next.properties.change
        ? revisionVisibility(next.properties.change.kind, mode)
        : "plain";
      if (nextVisibility !== "hidden") break;
    }
    if (nextIndex > index + 1) {
      output.push({ ...current, runs });
      index = nextIndex - 1;
    } else output.push(block);
  }
  return output;
};

const resolveDecoration = (blocks: readonly Block[] | undefined, mode: ChangesMode) => {
  return blocks ? resolveParagraphMarks(blocks, mode) : undefined;
};

export const resolveDocumentParagraphMarks = (
  doc: DocxDocument,
  mode: ChangesMode,
): DocxDocument => {
  const resolveProperties = (properties: SectionProperties): SectionProperties => ({
    ...properties,
    header: resolveDecoration(properties.header, mode),
    firstHeader: resolveDecoration(properties.firstHeader, mode),
    evenHeader: resolveDecoration(properties.evenHeader, mode),
    footer: resolveDecoration(properties.footer, mode),
    firstFooter: resolveDecoration(properties.firstFooter, mode),
    evenFooter: resolveDecoration(properties.evenFooter, mode),
  });
  return {
    ...doc,
    sections: doc.sections.map((section) => ({
      ...section,
      blocks: resolveParagraphMarks(section.blocks, mode),
      properties: resolveProperties(section.properties),
    })),
  };
};
