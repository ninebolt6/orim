import type { Inline } from "../docx/model";
import type { RenderContext } from "./context";
import { safeHref } from "./sanitize";
import { applyTextStyle, changeColor } from "./text-style";

export function renderInline(inline: Inline, context: RenderContext): Node {
  if (inline.type === "break") {
    return inline.page ? document.createTextNode("") : document.createElement("br");
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
    const change = inline.change;
    const hidden =
      change &&
      ((context.changes === "final" && (change.kind === "del" || change.kind === "moveFrom")) ||
        (context.changes === "original" && (change.kind === "ins" || change.kind === "moveTo")));
    if (hidden) return document.createTextNode("");
    const image = document.createElement("img");
    image.className = "orim-image";
    if (change) {
      image.classList.add(`orim-change--${change.kind}`);
      if (context.changes === "markup" && change.kind !== "moveTo" && change.kind !== "moveFrom")
        image.style.setProperty("--orim-change-color", changeColor(change, context.colors));
    }
    const part = context.doc.imageParts[inline.path];
    if (part && typeof URL.createObjectURL === "function") {
      let url = context.objectUrlsByPart.get(inline.path);
      if (!url) {
        const blob = new Blob([part.bytes.slice().buffer as ArrayBuffer], { type: part.mime });
        url = URL.createObjectURL(blob);
        context.objectUrlsByPart.set(inline.path, url);
        context.objectUrls.add(url);
      }
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
    applyTextStyle(field, inline.style);
    return field;
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
  } else textNode = document.createElement("span");
  textNode.classList.add("orim-text");
  textNode.textContent = inline.text;
  applyTextStyle(textNode, inline.style);
  const titles: string[] = [];
  if (inline.change) {
    const { kind, author, date } = inline.change;
    textNode.classList.add(`orim-change--${kind}`);
    if (kind === "moveTo" || kind === "moveFrom") {
      const accepted =
        (kind === "moveTo" && context.changes === "final") ||
        (kind === "moveFrom" && context.changes === "original");
      textNode.style.color = accepted ? "inherit" : "#137333";
      textNode.style.textDecoration = accepted
        ? "none"
        : kind === "moveTo"
          ? "underline double"
          : "line-through double";
    } else
      textNode.style.setProperty("--orim-change-color", changeColor(inline.change, context.colors));
    titles.push(
      `${kind === "ins" || kind === "moveTo" ? "Inserted" : "Deleted"} by ${author}${date ? ` on ${date}` : ""}`,
    );
  }
  if (inline.formattingChange && context.changes === "markup") {
    textNode.classList.add("orim-format-change");
    titles.push(
      `Formatted: ${inline.formattingChange.author}${inline.formattingChange.date ? ` ${inline.formattingChange.date}` : ""}`,
    );
  }
  if (titles.length) textNode.title = titles.join("\n");
  if (context.comments && inline.comments?.length) {
    const anchor = document.createElement("span");
    anchor.className = "orim-comment-anchor";
    anchor.dataset.c = inline.comments.join(" ");
    anchor.append(textNode);
    return anchor;
  }
  return textNode;
}
