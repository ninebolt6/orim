import {
  attr,
  descendants,
  elements,
  first,
  openOpc,
  type OpcPackage,
  type Relationship,
  type ThemeFonts,
} from "../opc";
import type {
  Block,
  ChangeKind,
  ChangeMark,
  DocxComment,
  DocxDocument,
  Inline,
  Paragraph,
  ParagraphProperties,
  SectionProperties,
  Table,
  TableCell,
  TableRow,
  TextStyle,
} from "./model";
import { parseNumbering } from "./numbering";
import { border, paragraphProperties, parseRunProperties, parseStyles } from "./styles";

const twips = (value: string | undefined, fallback: number): number => {
  const n = Number(value);
  return Number.isFinite(n) ? n / 20 : fallback;
};

function localAttr(element: Element, name: string): string | undefined {
  for (const item of Array.from(element.attributes))
    if (item.localName.split(":").at(-1) === name) return item.value;
  return undefined;
}

function children(element: ParentNode): Element[] {
  return Array.from(element.childNodes).filter((node): node is Element => node.nodeType === 1);
}

function textValue(element: Element | undefined): string {
  return element?.textContent ?? "";
}

function color(value: string | undefined): string | undefined {
  return value && /^[\da-f]{6}$/i.test(value) ? `#${value}` : undefined;
}

function cssWidth(element: Element | undefined): string | undefined {
  const value = attr(element, "w");
  if (value === undefined) return undefined;
  const type = attr(element, "type") ?? "dxa";
  if (type === "pct") {
    const percentage = value.endsWith("%") ? Number(value.slice(0, -1)) : Number(value) / 50;
    return Number.isFinite(percentage) ? `${percentage}%` : undefined;
  }
  if (type !== "dxa") return undefined;
  const twipValue = Number(value);
  return Number.isFinite(twipValue) ? `${twipValue / 20}pt` : undefined;
}

function changeFrom(element: Element, kind: ChangeKind): ChangeMark {
  return {
    kind,
    author: attr(element, "author") ?? "Unknown",
    date: attr(element, "date") ?? "",
  };
}

function relationshipId(element: Element): string | undefined {
  return localAttr(element, "id") ?? localAttr(element, "embed") ?? localAttr(element, "link");
}

function imagePart(
  element: Element,
  rels: Map<string, Relationship>,
  source: string,
  opc: OpcPackage,
): string | undefined {
  const blip = descendants(element, "blip")[0];
  const id = blip && relationshipId(blip);
  const relation = id && rels.get(id);
  return relation && !relation.external ? opc.resolve(source, relation.target) : undefined;
}

interface ParseContext {
  opc: OpcPackage;
  source: string;
  rels: Map<string, Relationship>;
  styles: ReturnType<typeof parseStyles>;
  theme: ThemeFonts;
  numbering: ReturnType<typeof parseNumbering>;
  commentsOpen: Set<string>;
  images: Set<string>;
  fields: FieldState[];
  paragraphStyle?: string;
}

interface FieldState {
  instruction: string;
  stage: "instruction" | "result";
  style?: TextStyle;
}

function fieldName(instruction: string): string {
  return instruction.trim().match(/^\S+/)?.[0]?.toUpperCase() ?? "";
}

function fieldHref(instruction: string): string | undefined {
  if (fieldName(instruction) !== "HYPERLINK") return undefined;
  const bookmark = instruction.match(/\\l\s+(?:"([^"]+)"|(\S+))/i);
  if (bookmark) return `#${bookmark[1] ?? bookmark[2]}`;
  const target = instruction.match(/\bHYPERLINK\s+(?:"([^"]+)"|(\S+))/i);
  return target?.[1] ?? target?.[2];
}

function liveField(field: FieldState): boolean {
  const name = fieldName(field.instruction);
  return name === "PAGE" || name === "NUMPAGES";
}

function collectInline(
  parent: ParentNode,
  context: ParseContext,
  comments: Set<string>,
  inheritedChange?: ChangeMark,
): Inline[] {
  const out: Inline[] = [];
  const emit = (inline: Inline) => {
    const style = inline.type === "text" || inline.type === "field" ? inline.style : undefined;
    let suppressed = false;
    for (const field of context.fields) {
      if (liveField(field) && field.stage === "result" && !field.style && style)
        field.style = style;
      if (field.stage === "instruction" || (liveField(field) && field.stage === "result"))
        suppressed = true;
    }
    if (suppressed) return;
    const hyperlink = [...context.fields]
      .reverse()
      .find((field) => field.stage === "result" && fieldName(field.instruction) === "HYPERLINK");
    const href = hyperlink && fieldHref(hyperlink.instruction);
    out.push(href && inline.type === "text" ? { ...inline, href } : inline);
  };
  const handleFieldChar = (node: Element, style?: TextStyle) => {
    const type = attr(node, "fldCharType");
    if (type === "begin") context.fields.push({ instruction: "", stage: "instruction", style });
    else if (type === "separate") {
      const field = context.fields.at(-1);
      if (field) field.stage = "result";
    } else if (type === "end") {
      const field = context.fields.pop();
      if (field && liveField(field))
        emit({ type: "field", name: fieldName(field.instruction), style: field.style });
    }
  };

  for (const node of children(parent)) {
    const name = node.localName;
    if (name === "commentRangeStart") {
      const id = attr(node, "id");
      if (id) comments.add(id);
      continue;
    }
    if (name === "commentRangeEnd") {
      const id = attr(node, "id");
      if (id) comments.delete(id);
      continue;
    }
    if (name === "bookmarkStart") {
      const bookmark = attr(node, "name");
      if (bookmark) emit({ type: "bookmark", name: bookmark });
      continue;
    }
    if (name === "commentReference") {
      const id = attr(node, "id");
      if (id) emit({ type: "comment-reference", id });
      continue;
    }
    if (name === "fldSimple") {
      const instruction = attr(node, "instr") ?? "";
      const currentName = fieldName(instruction);
      if (currentName === "PAGE" || currentName === "NUMPAGES") {
        const state: FieldState = { instruction, stage: "result" };
        context.fields.push(state);
        collectInline(node, context, comments, inheritedChange);
        context.fields.pop();
        emit({ type: "field", name: currentName, style: state.style });
      } else {
        const nested = collectInline(node, context, comments, inheritedChange);
        const href = fieldHref(instruction);
        for (const item of nested) emit(href && item.type === "text" ? { ...item, href } : item);
      }
      continue;
    }
    if (name === "fldChar") {
      handleFieldChar(node);
      continue;
    }
    if (name === "instrText" && context.fields.at(-1)?.stage === "instruction") {
      context.fields.at(-1)!.instruction += textValue(node);
      continue;
    }
    if (name === "proofErr" || name === "object" || name === "pict") continue;
    if (name === "AlternateContent") {
      const fallback = elements(node, "Fallback")[0];
      if (fallback)
        for (const item of collectInline(fallback, context, comments, inheritedChange)) emit(item);
      continue;
    }
    if (name === "hyperlink") {
      const rel = relationshipId(node);
      const target = rel ? context.rels.get(rel) : undefined;
      const anchor = attr(node, "anchor");
      const href = anchor ? `#${anchor}` : target?.external ? target.target : undefined;
      for (const item of collectInline(node, context, comments, inheritedChange))
        emit(href && item.type === "text" ? { ...item, href } : item);
      continue;
    }
    const kind: ChangeKind | undefined =
      name === "ins"
        ? "ins"
        : name === "del"
          ? "del"
          : name === "moveFrom"
            ? "moveFrom"
            : name === "moveTo"
              ? "moveTo"
              : undefined;
    if (kind) {
      for (const item of collectInline(node, context, comments, changeFrom(node, kind))) emit(item);
      continue;
    }
    if (name === "drawing" || name === "anchor" || name === "inline") {
      const path = imagePart(node, context.rels, context.source, context.opc);
      if (path) {
        context.images.add(path);
        const extent = descendants(node, "extent")[0];
        const width = extent ? Number(localAttr(extent, "cx") ?? 0) / 12700 : undefined;
        const height = extent ? Number(localAttr(extent, "cy") ?? 0) / 12700 : undefined;
        emit({ type: "image", path, width, height });
      }
      continue;
    }
    if (name === "tab") {
      emit({ type: "text", text: "\t", comments: [...comments], change: inheritedChange });
      continue;
    }
    if (name === "br" || name === "cr") {
      emit({ type: "break", page: attr(node, "type") === "page" });
      continue;
    }
    if (name === "sym") {
      const code = attr(node, "char");
      if (code && /^[\da-f]{4}$/i.test(code))
        emit({
          type: "text",
          text: String.fromCharCode(parseInt(code, 16)),
          comments: [...comments],
          change: inheritedChange,
        });
      continue;
    }
    if (name === "r") {
      const rPr = first(node, "rPr");
      const styleId = attr(first(rPr ?? node, "rStyle"), "val");
      const base = context.styles.run(
        context.paragraphStyle,
        styleId,
        parseRunProperties(rPr, context.theme),
      );
      const rChange = first(rPr ?? node, "rPrChange");
      const formattingChange = rChange ? changeFrom(rChange, "ins") : undefined;
      for (const child of children(node)) {
        if (child.localName === "rPr") continue;
        if (child.localName === "fldChar") {
          handleFieldChar(child, base);
          continue;
        }
        if (child.localName === "instrText" && context.fields.at(-1)?.stage === "instruction") {
          context.fields.at(-1)!.instruction += textValue(child);
          continue;
        }
        if (["t", "delText", "delInstrText"].includes(child.localName)) {
          emit({
            type: "text",
            text: textValue(child),
            style: base,
            change: inheritedChange,
            formattingChange,
            comments: [...comments],
          });
        } else if (child.localName === "tab")
          emit({
            type: "text",
            text: "\t",
            style: base,
            change: inheritedChange,
            formattingChange,
            comments: [...comments],
          });
        else if (child.localName === "br" || child.localName === "cr")
          emit({ type: "break", page: attr(child, "type") === "page" });
        else if (child.localName === "noBreakHyphen")
          emit({
            type: "text",
            text: "\u2011",
            style: base,
            change: inheritedChange,
            formattingChange,
            comments: [...comments],
          });
        else if (child.localName === "softHyphen")
          emit({
            type: "text",
            text: "\u00ad",
            style: base,
            change: inheritedChange,
            formattingChange,
            comments: [...comments],
          });
        else if (child.localName === "drawing") {
          const nested = collectInline(
            { childNodes: [child] } as unknown as ParentNode,
            context,
            comments,
            inheritedChange,
          );
          for (const item of nested) emit(item);
        }
      }
      continue;
    }
    if (
      name === "sdt" ||
      name === "sdtContent" ||
      name === "smartTag" ||
      name === "customXml" ||
      name === "Fallback"
    ) {
      for (const item of collectInline(node, context, comments, inheritedChange)) emit(item);
    }
  }
  return out;
}

function parseParagraph(
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

function applyContextualSpacing(blocks: Block[]): Block[] {
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

function blockNodes(nodes: Element[]): Element[] {
  const out: Element[] = [];
  for (const node of nodes) {
    if (["p", "tbl", "sectPr"].includes(node.localName)) {
      out.push(node);
    } else if (node.localName === "AlternateContent") {
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

function parseTable(node: Element, context: ParseContext, comments: Set<string>): Table {
  const tblPr = first(node, "tblPr");
  const styleId = attr(first(tblPr ?? node, "tblStyle"), "val");
  const style = context.styles.table(styleId);
  const grid = first(node, "tblGrid");
  const widths = elements(grid ?? node, "gridCol").map((col) => cssWidth(col));
  const directMargins: NonNullable<Table["margins"]> = {};
  const directTableMargins = first(tblPr ?? node, "tblCellMar");
  for (const side of ["top", "right", "bottom", "left"] as const) {
    const item = first(directTableMargins ?? node, side);
    const value = item ? twips(attr(item, "w"), 0) : undefined;
    if (value !== undefined) directMargins[side] = value;
  }
  const tableMargins = { ...style.margins, ...directMargins };
  const rows: TableRow[] = elements(node, "tr").map((row) => {
    const trPr = first(row, "trPr");
    const cells: TableCell[] = elements(row, "tc").map((cell) => {
      const tcPr = first(cell, "tcPr");
      const span = Number(attr(first(tcPr ?? cell, "gridSpan"), "val") ?? 1);
      const merge = first(tcPr ?? cell, "vMerge");
      const val = attr(merge, "val");
      const shading = attr(first(tcPr ?? cell, "shd"), "fill");
      const tcW = first(tcPr ?? cell, "tcW");
      const cellMargins = first(tcPr ?? cell, "tcMar");
      const borders: NonNullable<TableCell["borders"]> = {};
      const cellBorders = first(tcPr ?? cell, "tcBorders");
      for (const side of ["top", "right", "bottom", "left"] as const) {
        const value = border(first(cellBorders ?? cell, side));
        if (value !== undefined) borders[side] = value;
      }
      const blocks = applyContextualSpacing(
        blockNodes(children(cell)).flatMap((child) =>
          child.localName === "p"
            ? splitAtPageBreaks(parseParagraph(child, context, comments))
            : child.localName === "tbl"
              ? [parseTable(child, context, comments)]
              : [],
        ),
      );
      const marginValue = (side: string): number | undefined => {
        const item = first(cellMargins ?? cell, side);
        return item ? twips(attr(item, "w"), 0) : tableMargins[side as keyof typeof tableMargins];
      };
      return {
        blocks,
        gridSpan: Math.max(1, span),
        rowSpan: val === "restart" ? -1 : val === "continue" || merge ? 0 : 1,
        width: cssWidth(tcW),
        shading: (shading && shading !== "auto" ? color(shading) : undefined) ?? style.cellShading,
        borders,
        margins: {
          top: marginValue("top"),
          right: marginValue("right"),
          bottom: marginValue("bottom"),
          left: marginValue("left"),
        },
        verticalAlign: attr(first(tcPr ?? cell, "vAlign"), "val"),
      };
    });
    return {
      cells,
      header: !!first(trPr ?? row, "tblHeader"),
      height: twips(attr(first(trPr ?? row, "trHeight"), "val"), 0),
      change: first(trPr ?? row, "ins")
        ? changeFrom(first(trPr ?? row, "ins")!, "ins")
        : first(trPr ?? row, "del")
          ? changeFrom(first(trPr ?? row, "del")!, "del")
          : undefined,
    };
  });

  const gridStarts = rows.map((row) => {
    let column = 0;
    return row.cells.map((cell) => {
      const start = column;
      column += cell.gridSpan;
      return start;
    });
  });
  for (let row = 0; row < rows.length; row++) {
    for (let cell = 0; cell < rows[row].cells.length; cell++) {
      const item = rows[row].cells[cell];
      if (item.rowSpan !== -1) continue;
      let span = 1;
      const start = gridStarts[row][cell];
      for (let next = row + 1; next < rows.length; next++) {
        const index = gridStarts[next].findIndex((column, candidate) => {
          const continuation = rows[next].cells[candidate];
          return (
            column === start &&
            continuation.rowSpan === 0 &&
            continuation.gridSpan === item.gridSpan
          );
        });
        if (index < 0) break;
        const continuation = rows[next].cells[index];
        continuation.rowSpan = -2;
        span++;
      }
      item.rowSpan = span;
    }
  }
  for (const row of rows) for (const cell of row.cells) if (cell.rowSpan === -2) cell.rowSpan = 0;

  const tableBorders: NonNullable<Table["borders"]> = { ...style.borders };
  const tblBorders = first(tblPr ?? node, "tblBorders");
  for (const side of ["top", "right", "bottom", "left", "insideH", "insideV"] as const) {
    const value = border(first(tblBorders ?? node, side));
    if (value !== undefined) tableBorders[side] = value;
  }
  return {
    type: "table",
    rows,
    widths,
    width: cssWidth(first(tblPr ?? node, "tblW")),
    indent: twips(attr(first(tblPr ?? node, "tblInd"), "w"), 0),
    align: attr(first(tblPr ?? node, "jc"), "val"),
    borders: tableBorders,
    margins: tableMargins,
  };
}

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

async function parseBlocks(
  nodes: Element[],
  context: ParseContext,
  comments: Set<string>,
): Promise<Block[]> {
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
      fields: [],
    };
    const blocks = await parseBlocks(children(root), partContext, new Set());
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
    fields: [],
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
      fields: [],
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

  const commentList: DocxComment[] = [];
  if (opc.zip.has("word/comments.xml")) {
    const commentsPath = "word/comments.xml";
    const commentsDoc = await opc.xml(commentsPath);
    const commentContext: ParseContext = {
      ...context,
      source: commentsPath,
      rels: await opc.relationships(commentsPath),
      commentsOpen: new Set(),
      fields: [],
    };
    for (const node of descendants(commentsDoc, "comment")) {
      const id = attr(node, "id");
      if (!id) continue;
      const blocks = await parseBlocks(children(node), commentContext, new Set());
      const lastParagraph = elements(node, "p").at(-1);
      const paraId = lastParagraph && localAttr(lastParagraph, "paraId");
      commentList.push({
        id,
        author: attr(node, "author") ?? "Unknown",
        initials: attr(node, "initials"),
        date: attr(node, "date"),
        paragraphs: blocks,
        paraId,
      });
    }
  }
  if (opc.zip.has("word/commentsExtended.xml")) {
    const extended = await opc.xml("word/commentsExtended.xml");
    const byPara = new Map(commentList.map((comment) => [comment.paraId ?? "", comment]));
    for (const node of descendants(extended, "commentEx")) {
      const paraId = localAttr(node, "paraId");
      const parent = localAttr(node, "paraIdParent");
      const comment = byPara.get(paraId ?? "");
      if (!comment) continue;
      comment.parentId = byPara.get(parent ?? "")?.id;
      comment.done = localAttr(node, "done") === "1" || localAttr(node, "done") === "true";
    }
  }

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
