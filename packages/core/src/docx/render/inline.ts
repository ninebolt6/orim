import type { Inline } from "../model";
import type { RenderContext } from "./context";
import { safeHref } from "../../common/sanitize";
import { applyTextStyle, changeColor } from "./text-style";
import { revisionVisibility } from "./revisions";

export const renderInline = (inline: Inline, context: RenderContext): Node => {
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
    let number = context.commentReferences.get(inline.id);
    if (number === undefined) {
      number = context.commentReferences.size + 1;
      context.commentReferences.set(inline.id, number);
    }
    marker.textContent = String(number);
    marker.dataset.comment = inline.id;
    return marker;
  }
  if (inline.type === "image") {
    const change = inline.change;
    const visibility = change ? revisionVisibility(change.kind, context.changes) : "plain";
    if (visibility === "hidden") return document.createTextNode("");
    const image = document.createElement("img");
    image.className = "orim-image";
    if (change && visibility === "marked") {
      image.classList.add(`orim-change--${change.kind}`);
      if (change.kind !== "moveTo" && change.kind !== "moveFrom")
        image.style.setProperty("--orim-change-color", changeColor(change, context.colors));
    }
    const url = context.objectUrls.url(inline.path, context.doc.imageParts[inline.path]);
    if (url) image.src = url;
    if (inline.width) image.style.width = `${inline.width}pt`;
    if (inline.height) image.style.height = `${inline.height}pt`;
    image.alt = "";
    return image;
  }
  if (inline.type === "field") {
    const field = document.createElement("span");
    field.className = "orim-field";
    field.textContent = inline.cached ?? "";
    applyTextStyle(field, inline.style);
    if (inline.name === "PAGE" && context.decorationPageNumber !== undefined)
      field.textContent = String(context.decorationPageNumber);
    else if (inline.name === "PAGE" || inline.name === "NUMPAGES")
      context.fields.push({
        element: field,
        name: inline.name,
        cached: inline.cached ?? "",
      });
    return field;
  }
  const visibility = inline.change
    ? revisionVisibility(inline.change.kind, context.changes)
    : "plain";
  if (visibility === "hidden") return document.createTextNode("");
  let textNode: HTMLElement;
  const href = safeHref(inline.href);
  if (href) {
    const link = document.createElement("a");
    link.className = "orim-link";
    link.href = href;
    link.target = href.startsWith("#") ? "" : "_blank";
    if (link.target) link.rel = "noopener noreferrer";
    textNode = link;
  } else textNode = document.createElement("span");
  textNode.classList.add("orim-text");
  textNode.textContent = inline.text;
  applyTextStyle(textNode, inline.style, visibility === "marked");
  const titles: string[] = [];
  if (inline.change) {
    const { kind, author, date } = inline.change;
    if (visibility === "marked") {
      textNode.classList.add(`orim-change--${kind}`);
      if (kind !== "moveTo" && kind !== "moveFrom")
        textNode.style.setProperty(
          "--orim-change-color",
          changeColor(inline.change, context.colors),
        );
      titles.push(
        `${kind === "ins" || kind === "moveTo" ? "Inserted" : "Deleted"} by ${author}${date ? ` on ${date}` : ""}`,
      );
    }
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
    for (const id of inline.comments) {
      const anchors = context.commentAnchors.get(id) ?? [];
      anchors.push(anchor);
      context.commentAnchors.set(id, anchors);
    }
    return anchor;
  }
  return textNode;
};
