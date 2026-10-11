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
} from "./model";
import { parseNumbering } from "./numbering";
import { parseRunProperties, parseStyles } from "./styles";

const twips = (value: string | undefined, fallback: number): number => {
  const n = Number(value);
  return Number.isFinite(n) ? n / 20 : fallback;
};

const number = (element: Element | undefined, name: string, fallback = 0): number => {
  const value = attr(element, name);
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
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
  paragraphStyle?: string;
}

interface FieldState {
  instruction: string;
  result: Inline[];
  stage: "instruction" | "result";
}

function plain(inlines: Inline[]): string {
  return inlines
    .map((inline) =>
      inline.type === "text" ? inline.text : inline.type === "field" ? (inline.cached ?? "") : "",
    )
    .join("");
}

function parseParagraphProperties(pPr: Element | undefined): ParagraphProperties {
  const spacing = first(pPr ?? document, "spacing");
  const indent = first(pPr ?? document, "ind");
  const jc = attr(first(pPr ?? document, "jc"), "val");
  const shading = first(pPr ?? document, "shd");
  const fill = attr(shading, "fill");
  const rule = attr(spacing, "lineRule");
  const line = attr(spacing, "line");
  const borders = first(pPr ?? document, "pBdr");
  const border = borders && (first(borders, "left") ?? first(borders, "bottom"));
  const borderColor = color(attr(border, "color"));
  const props: ParagraphProperties = {
    align: jc,
    before: twips(attr(spacing, "before"), 0),
    after: twips(attr(spacing, "after"), 0),
    line: rule === "auto" ? number(spacing, "line") / 240 : twips(line, 0),
    lineRule: rule,
    left: twips(attr(indent, "left") ?? attr(indent, "start"), 0),
    right: twips(attr(indent, "right") ?? attr(indent, "end"), 0),
    firstLine: twips(attr(indent, "firstLine"), 0),
    hanging: twips(attr(indent, "hanging"), 0),
    contextualSpacing: !!first(pPr ?? document, "contextualSpacing"),
    pageBreakBefore: !!first(pPr ?? document, "pageBreakBefore"),
    keepNext: !!first(pPr ?? document, "keepNext"),
    shading: color(fill),
    border: border
      ? `${Number(attr(border, "sz") ?? 4) / 8}pt solid ${borderColor ?? "#888"}`
      : undefined,
  };
  return props;
}

function collectInline(
  parent: ParentNode,
  context: ParseContext,
  comments: Set<string>,
  inheritedChange?: ChangeMark,
  field: { current?: FieldState } = {},
): Inline[] {
  const out: Inline[] = [];
  const emit = (inline: Inline) => {
    if (field.current?.stage === "result") field.current.result.push(inline);
    else if (!field.current || field.current.stage !== "instruction") out.push(inline);
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
      const nested = collectInline(node, context, comments, inheritedChange, {});
      emit({
        type: "field",
        name: (attr(node, "instr") ?? "").trim().split(/\s+/)[0]?.toUpperCase() ?? "",
        cached: plain(nested),
      });
      continue;
    }
    if (name === "fldChar") {
      const type = attr(node, "fldCharType");
      if (type === "begin") field.current = { instruction: "", result: [], stage: "instruction" };
      else if (type === "separate" && field.current) field.current.stage = "result";
      else if (type === "end" && field.current) {
        const completed = field.current;
        const instruction = completed.instruction.trim().split(/\s+/);
        const name = instruction[0]?.toUpperCase() ?? "";
        const cached = plain(completed.result);
        field.current = undefined;
        emit({ type: "field", name, cached });
      }
      continue;
    }
    if (name === "instrText" && field.current?.stage === "instruction") {
      field.current.instruction += textValue(node);
      continue;
    }
    if (name === "proofErr" || name === "object" || name === "pict") continue;
    if (name === "AlternateContent") {
      const fallback = descendants(node, "Fallback")[0];
      if (fallback) out.push(...collectInline(fallback, context, comments, inheritedChange, field));
      continue;
    }
    if (name === "hyperlink") {
      const rel = relationshipId(node);
      const target = rel ? context.rels.get(rel) : undefined;
      const anchor = attr(node, "anchor");
      const href = anchor ? `#${anchor}` : target?.external ? target.target : undefined;
      const nested = collectInline(node, context, comments, inheritedChange, field);
      out.push(...nested.map((item) => (item.type === "text" ? { ...item, href } : item)));
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
      out.push(...collectInline(node, context, comments, changeFrom(node, kind), field));
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
      const change = inheritedChange;
      const formattingChange = rChange ? changeFrom(rChange, "ins") : undefined;
      for (const child of children(node)) {
        if (child.localName === "rPr") continue;
        if (child.localName === "fldChar") {
          const type = attr(child, "fldCharType");
          if (type === "begin")
            field.current = { instruction: "", result: [], stage: "instruction" };
          else if (type === "separate" && field.current) field.current.stage = "result";
          else if (type === "end" && field.current) {
            const completed = field.current;
            const name = completed.instruction.trim().split(/\s+/)[0]?.toUpperCase() ?? "";
            field.current = undefined;
            emit({ type: "field", name, cached: plain(completed.result) });
          }
          continue;
        }
        if (child.localName === "instrText" && field.current?.stage === "instruction") {
          field.current.instruction += textValue(child);
          continue;
        }
        if (["t", "delText", "delInstrText"].includes(child.localName)) {
          emit({
            type: "text",
            text: textValue(child),
            style: base,
            change,
            formattingChange,
            comments: [...comments],
          });
        } else if (child.localName === "tab")
          emit({
            type: "text",
            text: "\t",
            style: base,
            change,
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
            change,
            formattingChange,
            comments: [...comments],
          });
        else if (child.localName === "softHyphen")
          emit({
            type: "text",
            text: "\u00ad",
            style: base,
            change,
            formattingChange,
            comments: [...comments],
          });
        else if (child.localName === "drawing") {
          const nested = collectInline(
            { childNodes: [child] } as unknown as ParentNode,
            context,
            comments,
            change,
            field,
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
      out.push(...collectInline(node, context, comments, inheritedChange, field));
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
  const properties = context.styles.paragraph(pStyle, parseParagraphProperties(pPr));
  properties.styleId = pStyle;
  const numPr = first(pPr ?? node, "numPr");
  const numId = attr(first(numPr ?? node, "numId"), "val");
  const level = Number(attr(first(numPr ?? node, "ilvl"), "val") ?? 0);
  const runs = collectInline(node, { ...context, paragraphStyle: pStyle }, commentState);
  const marker = numId ? context.numbering.marker(numId, level) : undefined;
  const markChange = first(pPr ?? node, "ins") ?? first(pPr ?? node, "del");
  if (markChange)
    properties.change = changeFrom(markChange, first(pPr ?? node, "del") ? "del" : "ins");
  const formattingChange = first(pPr ?? node, "pPrChange");
  if (formattingChange) properties.formattingChange = changeFrom(formattingChange, "ins");
  return {
    type: "paragraph",
    properties,
    runs,
    list: marker
      ? { numId: numId!, level, marker: marker.text, left: marker.left, hanging: marker.hanging }
      : undefined,
  };
}

function parseTable(node: Element, context: ParseContext, comments: Set<string>): Table {
  const tblPr = first(node, "tblPr");
  const grid = first(node, "tblGrid");
  const widths = elements(grid ?? node, "gridCol").map((col) => twips(attr(col, "w"), 0));
  const rows: TableRow[] = elements(node, "tr").map((row) => {
    const trPr = first(row, "trPr");
    const cells: TableCell[] = elements(row, "tc").map((cell) => {
      const tcPr = first(cell, "tcPr");
      const span = Number(attr(first(tcPr ?? cell, "gridSpan"), "val") ?? 1);
      const merge = first(tcPr ?? cell, "vMerge");
      const val = attr(merge, "val");
      const shading = attr(first(tcPr ?? cell, "shd"), "fill");
      const tcW = first(tcPr ?? cell, "tcW");
      const margins = first(tcPr ?? cell, "tcMar");
      const borders: Record<string, string> = {};
      for (const side of ["top", "right", "bottom", "left"]) {
        const b = first(first(tcPr ?? cell, "tcBorders") ?? cell, side);
        const c = color(attr(b, "color"));
        if (b && c) borders[side] = `${Number(attr(b, "sz") ?? 4) / 8}pt solid ${c}`;
      }
      const blocks = children(cell)
        .filter((child) => child.localName === "p" || child.localName === "tbl")
        .map((child) =>
          child.localName === "p"
            ? parseParagraph(child, context, comments)
            : parseTable(child, context, comments),
        );
      const marginValue = (side: string): number | undefined => {
        const item = first(margins ?? cell, side);
        return item ? twips(attr(item, "w"), 0) : undefined;
      };
      return {
        blocks,
        gridSpan: Math.max(1, span),
        rowSpan: val === "restart" ? -1 : val === "continue" || merge ? 0 : 1,
        width: tcW ? twips(attr(tcW, "w"), 0) : undefined,
        shading: color(shading),
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

  for (let row = 0; row < rows.length; row++) {
    for (let cell = 0; cell < rows[row].cells.length; cell++) {
      const item = rows[row].cells[cell];
      if (item.rowSpan !== -1) continue;
      let span = 1;
      for (let next = row + 1; next < rows.length; next++) {
        const continuation = rows[next].cells[cell];
        if (!continuation || continuation.rowSpan !== 0) break;
        continuation.rowSpan = -2;
        span++;
      }
      item.rowSpan = span;
    }
  }
  for (const row of rows) for (const cell of row.cells) if (cell.rowSpan === -2) cell.rowSpan = 0;

  const tableBorders: Record<string, string> = {};
  const tblBorders = first(tblPr ?? node, "tblBorders");
  for (const side of ["top", "right", "bottom", "left", "insideH", "insideV"]) {
    const border = first(tblBorders ?? node, side);
    const c = color(attr(border, "color"));
    if (border && c) tableBorders[side] = `${Number(attr(border, "sz") ?? 4) / 8}pt solid ${c}`;
  }
  const cellMargins = first(tblPr ?? node, "tblCellMar");
  const margin = (side: string) => {
    const item = first(cellMargins ?? node, side);
    return item ? twips(attr(item, "w"), 0) : undefined;
  };
  return {
    type: "table",
    rows,
    widths,
    width: twips(attr(first(tblPr ?? node, "tblW"), "w"), 0) || undefined,
    indent: twips(attr(first(tblPr ?? node, "tblInd"), "w"), 0),
    borders: tableBorders,
    margins: {
      top: margin("top"),
      right: margin("right"),
      bottom: margin("bottom"),
      left: margin("left"),
    },
  };
}

function sectionProperties(node: Element | undefined): SectionProperties {
  const pageSize = first(node ?? document, "pgSz");
  const pageMargins = first(node ?? document, "pgMar");
  const orientation = attr(pageSize, "orient") === "landscape" ? "landscape" : "portrait";
  let width = twips(attr(pageSize, "w"), 595 * 20);
  let height = twips(attr(pageSize, "h"), 842 * 20);
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
  };
}

async function parseBlocks(
  nodes: Element[],
  context: ParseContext,
  comments: Set<string>,
): Promise<Block[]> {
  const out: Block[] = [];
  for (const node of nodes) {
    if (node.localName === "p") out.push(parseParagraph(node, context, comments));
    else if (node.localName === "tbl") out.push(parseTable(node, context, comments));
  }
  return out;
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
  };
  const body = first(doc.documentElement, "body");
  const sections: DocxDocument["sections"] = [];
  let blocks: Block[] = [];
  let secPr: Element | undefined;
  for (const node of children(body ?? doc)) {
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
      sections.push({ properties: props, blocks });
      blocks = [];
      secPr = undefined;
    } else if (node.localName === "p" || node.localName === "tbl") {
      blocks.push(
        node.localName === "p"
          ? parseParagraph(node, context, commentsOpen)
          : parseTable(node, context, commentsOpen),
      );
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
        sections.push({ properties: props, blocks });
        blocks = [];
      }
    }
  }
  if (blocks.length || !sections.length)
    sections.push({ properties: sectionProperties(secPr), blocks });

  for (let index = 0; index < sections.length; index++) {
    const section = sections[index];
    const props = section.properties as SectionProperties & { _refs?: Record<string, string> };
    await sectionDecorations(opc, props, documentPath, rels, {
      opc,
      styles,
      theme,
      numbering,
      images,
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
    const commentsDoc = await opc.xml("word/comments.xml");
    for (const node of descendants(commentsDoc, "comment")) {
      const id = attr(node, "id");
      if (!id) continue;
      const blocks = await parseBlocks(elements(node, "p"), context, new Set());
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
  return {
    sections,
    comments: commentList,
    evenAndOddHeaders: !!(settings && descendants(settings, "evenAndOddHeaders").length),
    imageParts,
  };
}
