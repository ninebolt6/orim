import css from "./style.css?inline";
import type {
  Block,
  ChangeMark,
  DocxComment,
  DocxDocument,
  Inline,
  Paragraph,
  Table,
  TextStyle,
} from "../docx/model";
import { parseDocxModel } from "../docx/parse";
import { paginate, type Page } from "./paginate";

export interface RenderOptions {
  changes?: "markup" | "final" | "original";
  comments?: boolean;
}

export interface RenderedDocument {
  destroy(): void;
  setChanges(mode: RenderOptions["changes"]): void;
}

const palette = ["#137333", "#c5221f", "#185abc", "#a142f4", "#b06000"];
const highlightColors: Record<string, string> = {
  black: "#000",
  blue: "#00f",
  cyan: "#0ff",
  darkBlue: "#00f",
  darkCyan: "#0aa",
  darkGray: "#555",
  darkGreen: "#080",
  darkMagenta: "#a0a",
  darkRed: "#a00",
  darkYellow: "#aa0",
  green: "#0f0",
  lightGray: "#aaa",
  magenta: "#f0f",
  red: "#f00",
  white: "#fff",
  yellow: "#ff0",
};

function validColor(value: string | undefined): string | undefined {
  return value && /^#[\da-f]{6}$/i.test(value) ? value : undefined;
}

function safeHref(value: string | undefined): string | undefined {
  if (!value) return undefined;
  if (value.startsWith("#") && /^#[\w.-]+$/.test(value)) return value;
  try {
    const url = new URL(value);
    return ["http:", "https:", "mailto:"].includes(url.protocol) ? value : undefined;
  } catch {
    return undefined;
  }
}

function changeColor(change: ChangeMark, colors: Map<string, string>): string {
  let value = colors.get(change.author);
  if (!value) {
    value = palette[colors.size % palette.length];
    colors.set(change.author, value);
  }
  return value;
}

function fontStack(name: string | undefined): string | undefined {
  if (!name) return undefined;
  const names = [
    ...new Set(
      name
        .split(",")
        .map((item) => item.replace(/["'`;{}<>]/g, "").trim())
        .filter(Boolean),
    ),
  ];
  return names.length
    ? `${names.map((item) => `"${item}"`).join(", ")}, "Noto Sans", serif`
    : undefined;
}

function applyTextStyle(element: HTMLElement, style: TextStyle | undefined): void {
  if (!style) return;
  if (style.bold !== undefined) element.style.fontWeight = style.bold ? "700" : "400";
  if (style.italic !== undefined) element.style.fontStyle = style.italic ? "italic" : "normal";
  if (style.underline) {
    element.style.textDecorationLine = "underline";
    element.style.textDecorationStyle =
      style.underline === "double" ||
      style.underline === "dotted" ||
      style.underline === "dash" ||
      style.underline === "wave"
        ? style.underline === "dash"
          ? "dashed"
          : style.underline
        : "solid";
  }
  if (style.strike)
    element.style.textDecorationLine = `${element.style.textDecorationLine} line-through`.trim();
  if (style.doubleStrike) element.style.textDecoration = "line-through double";
  const color = validColor(style.color);
  if (color) element.style.color = color;
  if (style.highlight)
    element.style.backgroundColor =
      highlightColors[style.highlight] ?? validColor(style.highlight) ?? "";
  else if (validColor(style.backgroundColor))
    element.style.backgroundColor = style.backgroundColor!;
  if (style.fontSize && Number.isFinite(style.fontSize))
    element.style.fontSize = `${style.fontSize}pt`;
  const family = fontStack(style.fontFamily);
  if (family) element.style.fontFamily = family;
  if (style.verticalAlign)
    element.style.verticalAlign = style.verticalAlign === "superscript" ? "super" : "sub";
  if (style.caps) element.style.textTransform = "uppercase";
  if (style.smallCaps) element.style.fontVariant = "small-caps";
  if (style.hidden) element.style.visibility = "hidden";
  for (const [key, value] of Object.entries(style.style ?? {})) {
    if (/^[a-z-]+$/i.test(key) && typeof value === "string") element.style.setProperty(key, value);
  }
}

interface RenderContext {
  doc: DocxDocument;
  comments: boolean;
  changes: string;
  objectUrls: Set<string>;
  colors: Map<string, string>;
}

function renderInline(inline: Inline, context: RenderContext): Node {
  if (inline.type === "break") {
    const element = document.createElement(inline.page ? "span" : "br");
    if (inline.page) element.className = "orim-page-break";
    return element;
  }
  if (inline.type === "bookmark") {
    const anchor = document.createElement("span");
    if (/^[\w.-]+$/.test(inline.name)) anchor.id = inline.name;
    return anchor;
  }
  if (inline.type === "comment-reference") {
    if (!context.comments) return document.createTextNode("");
    const marker = document.createElement("sup");
    marker.className = "orim-comment-reference";
    marker.textContent = String(Number(inline.id) + 1);
    marker.dataset.comment = inline.id;
    return marker;
  }
  if (inline.type === "image") {
    const image = document.createElement("img");
    image.className = "orim-image";
    const part = context.doc.imageParts[inline.path];
    if (part && typeof URL.createObjectURL === "function") {
      const blob = new Blob([part.bytes.slice().buffer as ArrayBuffer], { type: part.mime });
      const url = URL.createObjectURL(blob);
      context.objectUrls.add(url);
      image.src = url;
    }
    if (inline.width) image.style.width = `${inline.width}pt`;
    if (inline.height) image.style.height = `${inline.height}pt`;
    image.alt = "";
    return image;
  }
  if (inline.type === "field") {
    const field = document.createElement("span");
    field.className = "orim-field";
    field.dataset.field = inline.name;
    field.dataset.cached = inline.cached ?? "";
    field.textContent = inline.cached ?? "";
    return field;
  }
  const kind = inline.change?.kind;
  if (
    (context.changes === "final" && (kind === "del" || kind === "moveFrom")) ||
    (context.changes === "original" && (kind === "ins" || kind === "moveTo"))
  ) {
    return document.createTextNode("");
  }

  let textNode: HTMLElement;
  const href = safeHref(inline.href);
  if (href) {
    const link = document.createElement("a");
    link.className = "orim-link";
    link.href = href;
    link.target = href.startsWith("#") ? "" : "_blank";
    if (link.target) link.rel = "noopener";
    textNode = link;
  } else {
    textNode = document.createElement("span");
  }
  textNode.classList.add("orim-text");
  textNode.textContent = inline.text;
  applyTextStyle(textNode, inline.style);
  if (inline.text === "\t") textNode.classList.add("orim-tab");
  if (inline.change && context.changes === "markup") {
    const { kind, author, date } = inline.change;
    textNode.classList.add(`orim-change--${kind}`);
    textNode.style.setProperty("--orim-change-color", changeColor(inline.change, context.colors));
    textNode.title = `${kind === "ins" || kind === "moveTo" ? "Inserted" : "Deleted"} by ${author}${date ? ` on ${date}` : ""}`;
  }
  if (inline.formattingChange && context.changes === "markup") {
    textNode.classList.add("orim-format-change");
    textNode.title = `Formatted: ${inline.formattingChange.author}${inline.formattingChange.date ? ` ${inline.formattingChange.date}` : ""}`;
  }
  if (context.comments && inline.comments?.length) {
    const anchor = document.createElement("span");
    anchor.className = "orim-comment-anchor";
    anchor.dataset.c = inline.comments.join(" ");
    anchor.append(textNode);
    return anchor;
  }
  return textNode;
}

function renderParagraph(block: Paragraph, context: RenderContext): HTMLElement {
  const paragraph = document.createElement("p");
  paragraph.className = "orim-paragraph";
  const props = block.properties;
  if (props.styleId) paragraph.classList.add(`orim-style-${props.styleId.replace(/[^\w-]/g, "")}`);
  if (props.align) {
    const align = props.align;
    paragraph.style.textAlign = ["both", "distribute"].includes(align)
      ? "justify"
      : align === "start"
        ? "left"
        : align === "end"
          ? "right"
          : align;
  }
  if (props.before) paragraph.style.marginTop = `${props.before}pt`;
  if (props.after !== undefined) paragraph.style.marginBottom = `${props.after}pt`;
  if (props.line)
    paragraph.style.lineHeight = props.lineRule === "auto" ? String(props.line) : `${props.line}pt`;
  if (props.left) paragraph.style.marginLeft = `${props.left}pt`;
  if (props.right) paragraph.style.marginRight = `${props.right}pt`;
  if (props.firstLine) paragraph.style.textIndent = `${props.firstLine}pt`;
  if (props.hanging) paragraph.style.textIndent = `${-props.hanging}pt`;
  const shading = validColor(props.shading);
  if (shading) paragraph.style.backgroundColor = shading;
  if (props.border) paragraph.style.borderLeft = props.border;
  if (props.change && context.changes === "markup") {
    paragraph.classList.add("orim-paragraph--change");
    paragraph.title = `Formatted by ${props.change.author}${props.change.date ? ` on ${props.change.date}` : ""}`;
  }
  if (props.formattingChange && context.changes === "markup") {
    paragraph.classList.add("orim-format-change");
    paragraph.title = `Formatted: ${props.formattingChange.author}${props.formattingChange.date ? ` ${props.formattingChange.date}` : ""}`;
  }
  if (props.pageBreakBefore) paragraph.style.breakBefore = "page";
  if (block.list) {
    paragraph.classList.add("orim-list");
    paragraph.style.marginLeft = `${Math.max(0, block.list.left ?? 18)}pt`;
    paragraph.style.textIndent = `${-(block.list.hanging ?? 18)}pt`;
    const marker = document.createElement("span");
    marker.className = "orim-list-marker";
    marker.textContent = block.list.marker;
    paragraph.append(marker);
  }
  for (const inline of block.runs) paragraph.append(renderInline(inline, context));
  if (!block.runs.length) {
    const spacer = document.createElement("span");
    spacer.textContent = "\u200b";
    paragraph.append(spacer);
  }
  if (props.change && context.changes === "markup") {
    const mark = document.createElement("span");
    mark.className = "orim-paragraph-mark";
    mark.textContent = " ¶";
    paragraph.append(mark);
  }
  if (props.contextualSpacing) paragraph.classList.add("orim-contextual-spacing");
  return paragraph;
}

function renderTable(block: Table, context: RenderContext): HTMLTableElement {
  const table = document.createElement("table");
  table.className = "orim-table";
  table.style.width = block.width ? `${block.width}pt` : "100%";
  if (block.indent) table.style.marginLeft = `${block.indent}pt`;
  const colgroup = document.createElement("colgroup");
  for (const width of block.widths) {
    const col = document.createElement("col");
    if (width) col.style.width = `${width}pt`;
    colgroup.append(col);
  }
  if (block.widths.length) table.append(colgroup);
  const body = document.createElement("tbody");
  for (const [rowIndex, row] of block.rows.entries()) {
    if (row.change?.kind === "del" && context.changes === "final") continue;
    if (row.change?.kind === "ins" && context.changes === "original") continue;
    const tr = document.createElement("tr");
    if (row.header) tr.dataset.orimHeader = "true";
    if (row.height) tr.style.height = `${row.height}pt`;
    if (row.change && context.changes === "markup") {
      tr.className = "orim-table-row--change";
      tr.title = `${row.change.kind === "ins" ? "Inserted" : "Deleted"} by ${row.change.author}`;
    }
    for (const [cellIndex, cell] of row.cells.entries()) {
      if (cell.rowSpan === 0) continue;
      const td = document.createElement("td");
      td.className = "orim-table-cell";
      if (cell.gridSpan > 1) td.colSpan = cell.gridSpan;
      if (cell.rowSpan > 1) td.rowSpan = cell.rowSpan;
      if (cell.width) td.style.width = `${cell.width}pt`;
      const bg = validColor(cell.shading);
      if (bg) td.style.backgroundColor = bg;
      if (cell.verticalAlign) td.style.verticalAlign = cell.verticalAlign;
      for (const side of ["top", "right", "bottom", "left"] as const) {
        const tableSide =
          side === "top"
            ? rowIndex === 0
              ? "top"
              : "insideH"
            : side === "bottom"
              ? rowIndex === block.rows.length - 1
                ? "bottom"
                : "insideH"
              : side === "left"
                ? cellIndex === 0
                  ? "left"
                  : "insideV"
                : cellIndex === row.cells.length - 1
                  ? "right"
                  : "insideV";
        const border = cell.borders?.[side] ?? block.borders?.[tableSide];
        if (border) td.style.setProperty(`border-${side}`, border);
        const margin =
          cell.margins?.[side] ??
          block.margins?.[side] ??
          (side === "left" || side === "right" ? 5.4 : 0);
        td.style.setProperty(`padding-${side}`, `${margin}pt`);
      }
      for (const nested of cell.blocks) td.append(renderBlock(nested, context));
      tr.append(td);
    }
    body.append(tr);
  }
  table.append(body);
  return table;
}

function renderBlock(block: Block, context: RenderContext): HTMLElement {
  if (block.type === "paragraph") return renderParagraph(block, context);
  return renderTable(block, context);
}

function renderCommentText(comment: DocxComment): string {
  return comment.paragraphs
    .map((block) =>
      block.type === "paragraph"
        ? block.runs.map((run) => (run.type === "text" ? run.text : "")).join("")
        : "",
    )
    .join("\n");
}

function addComments(pages: Page[], comments: DocxComment[]): void {
  const scope = pages[0]?.wrapper.parentElement;
  if (!scope) return;
  const byId = new Map(comments.map((comment) => [comment.id, comment]));
  for (const page of pages) {
    const lane = page.wrapper.querySelector<HTMLElement>(".orim-page-comments");
    if (!lane) continue;
    const candidates = Array.from(
      page.content.querySelectorAll<HTMLElement>(".orim-comment-anchor[data-c]"),
    );
    const ordered = new Map<string, number>();
    for (const anchor of candidates) {
      const top = anchor.getBoundingClientRect().top - page.page.getBoundingClientRect().top;
      for (const id of anchor.dataset.c?.split(" ") ?? []) {
        if (!ordered.has(id)) ordered.set(id, top);
      }
    }
    let nextTop = 0;
    for (const [id, top] of [...ordered].sort((a, b) => a[1] - b[1])) {
      const comment = byId.get(id);
      if (!comment) continue;
      const balloon = document.createElement("div");
      balloon.className = "orim-comment";
      balloon.dataset.comment = id;
      balloon.style.top = `${Math.max(top, nextTop)}px`;
      if (comment.done) balloon.classList.add("orim-comment--done");
      const author = document.createElement("div");
      author.className = "orim-comment-author";
      author.textContent = comment.author;
      balloon.append(author);
      const text = document.createElement("div");
      text.className = "orim-comment-text";
      text.textContent = renderCommentText(comment);
      balloon.append(text);
      const replies = comments.filter((reply) => reply.parentId === comment.id);
      if (replies.length) {
        const replyBox = document.createElement("div");
        replyBox.className = "orim-comment-replies";
        for (const reply of replies) {
          const replyEl = document.createElement("div");
          replyEl.className = `orim-comment${reply.done ? " orim-comment--done" : ""}`;
          const replyAuthor = document.createElement("div");
          replyAuthor.className = "orim-comment-author";
          replyAuthor.textContent = reply.author;
          const replyText = document.createElement("div");
          replyText.textContent = renderCommentText(reply);
          replyEl.append(replyAuthor, replyText);
          replyBox.append(replyEl);
        }
        balloon.append(replyBox);
      }
      lane.append(balloon);
      nextTop = Number.parseFloat(balloon.style.top) + balloon.offsetHeight + 8;
    }
  }

  const highlight = (id: string | undefined, active: boolean) => {
    if (!id) return;
    for (const anchor of scope.querySelectorAll<HTMLElement>(".orim-comment-anchor[data-c]")) {
      if (anchor.dataset.c?.split(" ").includes(id)) {
        anchor.classList.toggle(
          "orim-comment-active",
          active || anchor.classList.contains("orim-comment-pinned"),
        );
      }
    }
    for (const balloon of scope.querySelectorAll<HTMLElement>(".orim-comment[data-comment]")) {
      if (balloon.dataset.comment === id) {
        balloon.classList.toggle(
          "orim-comment-active",
          active || balloon.classList.contains("orim-comment-pinned"),
        );
      }
    }
  };
  const togglePinned = (id: string | undefined) => {
    if (!id) return;
    const balloon = scope.querySelector<HTMLElement>(
      `.orim-comment[data-comment="${CSS.escape(id)}"]`,
    );
    const active = !balloon?.classList.contains("orim-comment-pinned");
    for (const element of scope.querySelectorAll<HTMLElement>(
      `.orim-comment[data-comment="${CSS.escape(id)}"],.orim-comment-anchor[data-c]`,
    )) {
      if (
        element.classList.contains("orim-comment-anchor") &&
        !element.dataset.c?.split(" ").includes(id)
      )
        continue;
      element.classList.toggle("orim-comment-pinned", active);
    }
    highlight(id, active);
  };
  for (const balloon of scope.querySelectorAll<HTMLElement>(".orim-comment[data-comment]")) {
    balloon.addEventListener("mouseenter", () => highlight(balloon.dataset.comment, true));
    balloon.addEventListener("mouseleave", () => highlight(balloon.dataset.comment, false));
    balloon.addEventListener("click", () => togglePinned(balloon.dataset.comment));
  }
  for (const anchor of scope.querySelectorAll<HTMLElement>(".orim-comment-anchor")) {
    const id = anchor.dataset.c?.split(" ")[0];
    anchor.addEventListener("mouseenter", () => highlight(id, true));
    anchor.addEventListener("mouseleave", () => highlight(id, false));
    anchor.addEventListener("click", () => togglePinned(id));
  }
}

function ensureStyles(): void {
  if (document.querySelector("style[data-orim]")) return;
  const style = document.createElement("style");
  style.dataset.orim = "true";
  style.textContent = css;
  document.head.append(style);
}

function isModel(input: unknown): input is DocxDocument {
  return !!input && typeof input === "object" && Array.isArray((input as DocxDocument).sections);
}

async function bytes(input: ArrayBuffer | Uint8Array | Blob): Promise<ArrayBuffer | Uint8Array> {
  return input instanceof Blob ? input.arrayBuffer() : input;
}

export async function renderDocx(
  input: ArrayBuffer | Uint8Array | Blob | DocxDocument,
  container: HTMLElement,
  options: RenderOptions = {},
): Promise<RenderedDocument> {
  ensureStyles();
  const doc = isModel(input) ? input : await parseDocxModel(await bytes(input));
  const root = document.createElement("div");
  const objectUrls = new Set<string>();
  const context: RenderContext = {
    doc,
    comments: options.comments !== false,
    changes: options.changes ?? "markup",
    objectUrls,
    colors: new Map(),
  };
  container.replaceChildren(root);
  let pages: Page[] = [];
  const draw = () => {
    root.className = `orim-root orim-${context.changes}`;
    pages = paginate(
      root,
      doc,
      context.changes,
      (block) => renderBlock(block, context),
      context.comments,
    );
    if (context.comments) addComments(pages, doc.comments);
  };
  draw();
  return {
    destroy() {
      container.replaceChildren();
      for (const url of objectUrls) URL.revokeObjectURL(url);
      objectUrls.clear();
    },
    setChanges(mode) {
      if (!mode || mode === context.changes) return;
      context.changes = mode;
      draw();
    },
  };
}

export async function parseDocx(input: ArrayBuffer | Uint8Array | Blob): Promise<DocxDocument> {
  return parseDocxModel(await bytes(input));
}
