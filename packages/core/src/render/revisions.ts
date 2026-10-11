import type { Block, DocxDocument, Paragraph, SectionProperties, Table } from "../docx/model";
import type { ChangesMode } from "./document";

function resolveTable(table: Table, mode: ChangesMode): Table {
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
}

export function resolveParagraphMarks(blocks: readonly Block[], mode: ChangesMode): Block[] {
  const resolved = blocks.map((block) =>
    block.type === "table" ? resolveTable(block, mode) : block,
  );
  if (mode === "markup") return resolved;

  const excludedKind = mode === "final" ? "del" : "ins";
  const output: Block[] = [];
  for (let index = 0; index < resolved.length; index++) {
    const block = resolved[index];
    if (block.type !== "paragraph" || block.properties.change?.kind !== excludedKind) {
      output.push(block);
      continue;
    }

    let current: Paragraph = block;
    let runs = [...block.runs];
    let nextIndex = index + 1;
    while (resolved[nextIndex]?.type === "paragraph") {
      const next = resolved[nextIndex] as Paragraph;
      runs = [...runs, ...next.runs];
      current = next;
      nextIndex++;
      if (current.properties.change?.kind !== excludedKind) break;
    }
    if (nextIndex > index + 1) {
      output.push({ ...current, runs });
      index = nextIndex - 1;
    } else output.push(block);
  }
  return output;
}

function resolveDecoration(blocks: readonly Block[] | undefined, mode: ChangesMode) {
  return blocks ? resolveParagraphMarks(blocks, mode) : undefined;
}

export function resolveDocumentParagraphMarks(doc: DocxDocument, mode: ChangesMode): DocxDocument {
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
}
