import { attr, children, descendants, elements, first, textValue } from "../../xml";
import type { ChangeKind, ChangeMark, Inline, TextStyle } from "../model";
import type { OpcPackage, Relationship } from "../../opc/package";
import type { ParseContext } from "./context";
import { fieldHref, fieldName, isPageField, type FieldState } from "./fields";
import { parseRunProperties } from "../styles";

export function changeFrom(element: Element, kind: ChangeKind): ChangeMark {
  return {
    kind,
    author: attr(element, "author") ?? "Unknown",
    date: attr(element, "date") ?? "",
  };
}

export function relationshipId(element: Element): string | undefined {
  return attr(element, "id") ?? attr(element, "embed") ?? attr(element, "link");
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

function liveField(field: FieldState): boolean {
  return isPageField(field);
}

export function collectInline(
  parent: ParentNode,
  context: ParseContext,
  inheritedChange?: ChangeMark,
): Inline[] {
  const out: Inline[] = [];
  const emit = (inline: Inline) => {
    const style = inline.type === "text" || inline.type === "field" ? inline.style : undefined;
    if (style) context.fields.captureStyle(style);
    const suppressed = context.fields.some(
      (field) => field.stage === "instruction" || (liveField(field) && field.stage === "result"),
    );
    if (suppressed) return;
    const hyperlink = context.fields.findLast(
      (field) => field.stage === "result" && fieldName(field.instruction) === "HYPERLINK",
    );
    const href = hyperlink && fieldHref(hyperlink.instruction);
    out.push(href && inline.type === "text" ? { ...inline, href } : inline);
  };
  const handleFieldChar = (node: Element, style?: TextStyle) => {
    const type = attr(node, "fldCharType");
    if (type === "begin") context.fields.begin(style);
    else if (type === "separate") context.fields.separate();
    else if (type === "end") {
      const field = context.fields.end();
      if (field && liveField(field))
        emit({ type: "field", name: fieldName(field.instruction), style: field.style });
    }
  };

  for (const node of children(parent)) {
    const name = node.localName;
    if (name === "commentRangeStart") {
      const id = attr(node, "id");
      if (id) context.commentRanges.start(id);
      continue;
    }
    if (name === "commentRangeEnd") {
      const id = attr(node, "id");
      if (id) context.commentRanges.end(id);
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
        context.fields.push({ instruction, stage: "result" });
        collectInline(node, context, inheritedChange);
        const field = context.fields.end();
        emit({ type: "field", name: currentName, style: field?.style });
      } else {
        const nested = collectInline(node, context, inheritedChange);
        const href = fieldHref(instruction);
        for (const item of nested) emit(href && item.type === "text" ? { ...item, href } : item);
      }
      continue;
    }
    if (name === "fldChar") {
      handleFieldChar(node);
      continue;
    }
    if (name === "instrText" && context.fields.current()?.stage === "instruction") {
      context.fields.append(textValue(node));
      continue;
    }
    if (name === "proofErr" || name === "object" || name === "pict") continue;
    if (name === "AlternateContent") {
      const fallback = elements(node, "Fallback")[0];
      if (fallback)
        for (const item of collectInline(fallback, context, inheritedChange)) emit(item);
      continue;
    }
    if (name === "hyperlink") {
      const rel = relationshipId(node);
      const target = rel ? context.rels.get(rel) : undefined;
      const anchor = attr(node, "anchor");
      const href = anchor ? `#${anchor}` : target?.external ? target.target : undefined;
      for (const item of collectInline(node, context, inheritedChange))
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
      for (const item of collectInline(node, context, changeFrom(node, kind))) emit(item);
      continue;
    }
    if (name === "drawing" || name === "anchor" || name === "inline") {
      const path = imagePart(node, context.rels, context.source, context.opc);
      if (path) {
        context.imageParts.add(path);
        const extent = descendants(node, "extent")[0];
        const width = extent ? Number(attr(extent, "cx") ?? 0) / 12700 : undefined;
        const height = extent ? Number(attr(extent, "cy") ?? 0) / 12700 : undefined;
        emit({ type: "image", path, width, height, change: inheritedChange });
      }
      continue;
    }
    if (name === "tab") {
      emit({
        type: "text",
        text: "\t",
        comments: context.commentRanges.snapshot(),
        change: inheritedChange,
      });
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
          comments: context.commentRanges.snapshot(),
          change: inheritedChange,
        });
      continue;
    }
    if (name === "r") {
      const rPr = first(node, "rPr");
      const styleId = attr(first(rPr, "rStyle"), "val");
      const base = context.styles.run(
        context.paragraphStyle,
        styleId,
        parseRunProperties(rPr, context.theme),
      );
      const rChange = first(rPr, "rPrChange");
      const formattingChange = rChange ? changeFrom(rChange, "ins") : undefined;
      for (const child of children(node)) {
        if (child.localName === "rPr") continue;
        if (child.localName === "commentReference") {
          const id = attr(child, "id");
          if (id) emit({ type: "comment-reference", id });
          continue;
        }
        if (child.localName === "fldChar") {
          handleFieldChar(child, base);
          continue;
        }
        if (child.localName === "instrText" && context.fields.current()?.stage === "instruction") {
          context.fields.append(textValue(child));
          continue;
        }
        if (["t", "delText", "delInstrText"].includes(child.localName)) {
          emit({
            type: "text",
            text: textValue(child),
            style: base,
            change: inheritedChange,
            formattingChange,
            comments: context.commentRanges.snapshot(),
          });
        } else if (child.localName === "tab")
          emit({
            type: "text",
            text: "\t",
            style: base,
            change: inheritedChange,
            formattingChange,
            comments: context.commentRanges.snapshot(),
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
            comments: context.commentRanges.snapshot(),
          });
        else if (child.localName === "softHyphen")
          emit({
            type: "text",
            text: "\u00ad",
            style: base,
            change: inheritedChange,
            formattingChange,
            comments: context.commentRanges.snapshot(),
          });
        else if (child.localName === "drawing") {
          const nested = collectInline(
            { childNodes: [child] } as unknown as ParentNode,
            context,
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
      for (const item of collectInline(node, context, inheritedChange)) emit(item);
    }
  }
  return out;
}
