import { attr, children, descendants, elements, first, R, textValue } from "../xml";
import type { ChangeKind, ChangeMark, Inline, RevisionInfo, TextStyle } from "../model";
import { blipPart } from "../../drawingml/blip";
import { emu } from "../../drawingml/units";
import type { ParseContext } from "./context";
import {
  fieldHref,
  fieldName,
  fieldStateHref,
  fieldStateName,
  isPageField,
  type FieldState,
} from "./fields";
import { parseRunProperties } from "../styles";

export const revisionInfoFrom = (element: Element): RevisionInfo => {
  return {
    author: attr(element, "author") ?? "Unknown",
    date: attr(element, "date") ?? "",
  };
};

export const changeFrom = (element: Element, kind: ChangeKind): ChangeMark => {
  return { ...revisionInfoFrom(element), kind };
};

export const relationshipId = (element: Element): string | undefined => {
  return attr(element, "id", R) ?? attr(element, "embed", R) ?? attr(element, "link", R);
};

const liveField = (field: FieldState): boolean => {
  return isPageField(field);
};

type EmitInline = (inline: Inline) => void;

const textInline = (
  node: Element,
  style: TextStyle | undefined,
  change: ChangeMark | undefined,
  formattingChange: RevisionInfo | undefined,
  context: ParseContext,
): Inline | undefined => {
  let text: string;
  switch (node.localName) {
    case "tab":
      text = "\t";
      break;
    case "sym": {
      const code = attr(node, "char");
      if (!code || !/^[\da-f]{4}$/i.test(code)) return undefined;
      text = String.fromCharCode(parseInt(code, 16));
      break;
    }
    case "noBreakHyphen":
      text = "\u2011";
      break;
    case "softHyphen":
      text = "\u00ad";
      break;
    case "t":
    case "delText":
    case "delInstrText":
      text = textValue(node);
      break;
    default:
      return undefined;
  }
  return {
    type: "text",
    text,
    style,
    change,
    formattingChange,
    comments: context.state.commentRanges.snapshot(),
  };
};

const collectNode = (
  node: Element,
  context: ParseContext,
  change: ChangeMark | undefined,
  emit: EmitInline,
  style?: TextStyle,
  formattingChange?: RevisionInfo,
): void => {
  const emitText = (element: Element) => {
    const inline = textInline(element, style, change, formattingChange, context);
    if (inline) emit(inline);
  };
  const handleFieldChar = (element: Element, fieldStyle?: TextStyle) => {
    const type = attr(element, "fldCharType");
    if (type === "begin") context.state.fields.begin(fieldStyle);
    else if (type === "separate") context.state.fields.separate();
    else if (type === "end") {
      const field = context.state.fields.end();
      if (field && liveField(field))
        emit({ type: "field", name: fieldStateName(field), style: field.style });
    }
  };
  const name = node.localName;
  if (name === "commentRangeStart" || name === "commentRangeEnd") {
    const id = attr(node, "id");
    if (id && name === "commentRangeStart") context.state.commentRanges.start(id);
    else if (id) context.state.commentRanges.end(id);
  } else if (name === "commentReference") {
    const id = attr(node, "id");
    if (id) emit({ type: "comment-reference", id });
  } else if (name === "bookmarkStart") {
    const bookmark = attr(node, "name");
    if (bookmark) emit({ type: "bookmark", name: bookmark });
  } else if (name === "fldChar") handleFieldChar(node, style);
  else if (name === "instrText") {
    if (context.state.fields.current()?.stage === "instruction")
      context.state.fields.append(textValue(node));
  } else if (name === "fldSimple") {
    const instruction = attr(node, "instr") ?? "";
    const currentName = fieldName(instruction);
    if (currentName === "PAGE" || currentName === "NUMPAGES") {
      context.state.fields.push({ instruction, stage: "result" });
      for (const child of children(node))
        collectNode(child, context, change, emit, style, formattingChange);
      const field = context.state.fields.end();
      emit({ type: "field", name: currentName, style: field?.style });
    } else {
      const href = fieldHref(instruction);
      const nestedEmit = (inline: Inline) =>
        emit(href && inline.type === "text" ? { ...inline, href } : inline);
      for (const child of children(node))
        collectNode(child, context, change, nestedEmit, style, formattingChange);
    }
  } else if (name === "r") {
    const rPr = first(node, "rPr");
    const base = context.config.styles.run(
      context.config.paragraphStyle,
      attr(first(rPr, "rStyle"), "val"),
      parseRunProperties(rPr, context.config.theme),
      node.textContent ?? "",
    );
    const rChange = first(rPr, "rPrChange");
    const runFormatting = rChange ? revisionInfoFrom(rChange) : undefined;
    for (const child of children(node))
      if (child.localName !== "rPr") collectNode(child, context, change, emit, base, runFormatting);
  } else if (name === "hyperlink") {
    const rel = relationshipId(node);
    const target = rel ? context.config.rels.get(rel) : undefined;
    const anchor = attr(node, "anchor");
    const href = anchor ? `#${anchor}` : target?.external ? target.target : undefined;
    const nestedEmit = (inline: Inline) =>
      emit(href && inline.type === "text" ? { ...inline, href } : inline);
    for (const child of children(node))
      collectNode(child, context, change, nestedEmit, style, formattingChange);
  } else {
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
      const nestedChange = changeFrom(node, kind);
      for (const child of children(node))
        collectNode(child, context, nestedChange, emit, style, formattingChange);
    } else if (name === "drawing" || name === "anchor" || name === "inline") {
      const path = blipPart(node, context.config.rels, context.config.source, context.config.opc);
      if (path) {
        const extent = descendants(node, "extent")[0];
        const width = extent ? (emu(attr(extent, "cx", null)) ?? 0) : undefined;
        const height = extent ? (emu(attr(extent, "cy", null)) ?? 0) : undefined;
        emit({ type: "image", path, width, height, change });
      }
    } else if (name === "br" || name === "cr")
      emit({ type: "break", page: attr(node, "type") === "page" });
    else if (
      ["t", "delText", "delInstrText", "tab", "sym", "noBreakHyphen", "softHyphen"].includes(name)
    )
      emitText(node);
    else if (name !== "proofErr" && name !== "object" && name !== "pict") {
      const fallback = name === "AlternateContent" ? elements(node, "Fallback")[0] : undefined;
      if (fallback) {
        for (const child of children(fallback))
          collectNode(child, context, change, emit, style, formattingChange);
      } else if (
        name === "AlternateContent" ||
        name === "sdt" ||
        name === "sdtContent" ||
        name === "smartTag" ||
        name === "customXml" ||
        name === "Fallback"
      ) {
        for (const child of children(node))
          collectNode(child, context, change, emit, style, formattingChange);
      }
    }
  }
};

export const collectInline = (
  parent: ParentNode,
  context: ParseContext,
  inheritedChange?: ChangeMark,
): Inline[] => {
  const out: Inline[] = [];
  const emit = (inline: Inline) => {
    const inlineStyle =
      inline.type === "text" || inline.type === "field" ? inline.style : undefined;
    if (inlineStyle) context.state.fields.captureStyle(inlineStyle);
    if (
      context.state.fields.some(
        (field) => field.stage === "instruction" || (liveField(field) && field.stage === "result"),
      )
    )
      return;
    const hyperlink = context.state.fields.findLast(
      (field) => field.stage === "result" && fieldStateName(field) === "HYPERLINK",
    );
    const href = hyperlink && fieldStateHref(hyperlink);
    out.push(href && inline.type === "text" ? { ...inline, href } : inline);
  };
  for (const node of children(parent)) collectNode(node, context, inheritedChange, emit);
  return out;
};
