import { attr, descendants, elements, first, R } from "../../common/xml";
import { themeColorScheme, themeFonts, themeFormats } from "../../drawingml/theme";
import { emu } from "../../drawingml/units";
import {
  loadOpc,
  optionalXml,
  readMedia,
  relatedPart,
  type OpcPackage,
  type PackageInput,
} from "../../opc/package";
import type { Fill, PptxPresentation, Slide, SlideElement } from "../model";
import type { PartContext, ShapeContext, SlideContext } from "./context";
import { placeholderOf, parseTree, type InheritedShape, type Placeholder } from "./shapes";
import { fillElement, parseFill } from "../../drawingml/fill";
import { fillOptions, referenceFill } from "./style";
import { parseTableStyles } from "./table";
import { parseListStyle, type ListStyle } from "./text";

interface Part extends PartContext {
  readonly root: Element;
}

const masterType = (type: string): string =>
  type === "ctrTitle"
    ? "title"
    : ["title", "dt", "ftr", "sldNum", "hdr"].includes(type)
      ? type
      : "body";

const placeholders = (part: Part | undefined): readonly [Element, Placeholder][] =>
  part
    ? ["sp", "pic", "graphicFrame"]
        .flatMap((name) => descendants(part.root, name))
        .flatMap((element) => {
          const placeholder = placeholderOf(element);
          return placeholder ? [[element, placeholder] as const] : [];
        })
    : [];

const matchPlaceholder = (
  candidates: readonly [Element, Placeholder][],
  key: Placeholder,
  byIndex: boolean,
): readonly [Element, Placeholder] | undefined =>
  (byIndex && key.index !== undefined
    ? candidates.find(([, item]) => item.index === key.index)
    : undefined) ??
  candidates.find(([, item]) => item.type === key.type) ??
  candidates.find(([, item]) => masterType(item.type) === masterType(key.type));

const colorMap = (
  master: Part,
  overrides: readonly (Part | undefined)[],
): Record<string, string> => {
  const map = (element: Element | undefined) =>
    Object.fromEntries(
      [...(element?.attributes ?? [])].map((item) => [item.localName, item.value]),
    );
  return overrides.reduce(
    (result, part) => {
      const override = first(first(part?.root, "clrMapOvr"), "overrideClrMapping");
      return override ? map(override) : result;
    },
    map(first(master.root, "clrMap")),
  );
};

const background = (parts: readonly Part[], context: SlideContext): Fill | undefined =>
  parts.reduceRight<Fill | undefined>((result, part) => {
    if (result) return result;
    const bg = first(first(part.root, "cSld"), "bg");
    const properties = first(bg, "bgPr");
    const shapeContext: ShapeContext = { ...context, part };
    return properties
      ? parseFill(fillElement(properties), fillOptions(shapeContext))
      : referenceFill(first(bg, "bgRef"), context);
  }, undefined);

const mediaParts = (elements: readonly SlideElement[]): string[] =>
  elements.flatMap((element): string[] => {
    switch (element.kind) {
      case "picture":
        return [element.part];
      case "group":
        return mediaParts(element.children);
      case "shape":
        return element.fill?.type === "image" ? [element.fill.part] : [];
      case "table":
        return element.rows.flatMap((row) =>
          row.cells.flatMap((cell) => (cell.fill?.type === "image" ? [cell.fill.part] : [])),
        );
    }
  });

export const parsePptxModel = async (input: PackageInput): Promise<PptxPresentation> => {
  const opc = await loadOpc(input);
  const presentationPath = await relatedPart(opc, "", "officeDocument");
  if (!presentationPath) throw new Error("Not a PPTX package: presentation part not found.");
  const presentation = (await opc.xml(presentationPath)).documentElement;
  if (presentation.localName !== "presentation")
    throw new Error("Not a PPTX package: presentation part is not a presentation.");
  const presentationRels = await opc.relationships(presentationPath);
  const size = first(presentation, "sldSz");
  const slidePaths = elements(first(presentation, "sldIdLst"), "sldId").flatMap((slide) => {
    const relation = presentationRels.get(attr(slide, "id", R) ?? "");
    return relation && !relation.external ? [opc.resolve(presentationPath, relation.target)] : [];
  });
  const tableStyles = parseTableStyles(
    await optionalXml(opc, await relatedPart(opc, presentationPath, "tableStyles")),
  );
  const parts = new Map<string, Promise<Part>>();
  const loadPart = (path: string): Promise<Part> => {
    let part = parts.get(path);
    if (!part) {
      part = readPart(opc, path);
      parts.set(path, part);
    }
    return part;
  };
  const themes = new Map<string, Promise<Document | undefined>>();
  const loadTheme = (master: string): Promise<Document | undefined> => {
    let theme = themes.get(master);
    if (!theme) {
      theme = relatedPart(opc, master, "theme").then((path) => optionalXml(opc, path));
      themes.set(master, theme);
    }
    return theme;
  };
  const slideNumbers = new Map(slidePaths.map((path, index) => [path, index + 1]));
  const slides = await Promise.all(
    slidePaths.map(async (path, index): Promise<Slide> => {
      const slide = await loadPart(path);
      if (slide.root.localName !== "sld") throw new Error(`Invalid PPTX slide part: ${path}`);
      const layoutPath = await relatedPart(opc, path, "slideLayout");
      const layout = layoutPath ? await loadPart(layoutPath) : undefined;
      const masterPath = layoutPath && (await relatedPart(opc, layoutPath, "slideMaster"));
      const master = masterPath ? await loadPart(masterPath) : undefined;
      const theme = masterPath ? await loadTheme(masterPath) : undefined;
      const colors = themeColorScheme(theme);
      const map = master ? colorMap(master, [layout, slide]) : {};
      const base: SlideContext = {
        scheme: (name) => colors[map[name] ?? name],
        fonts: themeFonts(theme),
        formats: themeFormats(theme),
        slideNumber: index + 1,
        defaultText: [],
        tableStyles,
        slidePaths: slideNumbers,
      };
      const context: SlideContext = {
        ...base,
        defaultText: parseListStyle(first(presentation, "defaultTextStyle"), base),
      };
      const textStyles = first(master?.root, "txStyles");
      const textStyle = (name: string): ListStyle =>
        parseListStyle(first(textStyles, name), context);
      const styles = {
        title: textStyle("titleStyle"),
        body: textStyle("bodyStyle"),
        other: textStyle("otherStyle"),
      };
      const layoutPlaceholders = placeholders(layout);
      const masterPlaceholders = placeholders(master);
      const inherit = (key: Placeholder) => {
        const fromLayout = matchPlaceholder(layoutPlaceholders, key, true);
        const fromMaster = matchPlaceholder(masterPlaceholders, fromLayout?.[1] ?? key, false);
        const type = masterType(key.type);
        return {
          sources: [
            ...(fromMaster && master ? [{ element: fromMaster[0], part: master }] : []),
            ...(fromLayout && layout ? [{ element: fromLayout[0], part: layout }] : []),
          ] satisfies InheritedShape[],
          textStyle: type === "title" ? styles.title : type === "body" ? styles.body : styles.other,
        };
      };
      const tree = (part: Part | undefined, template: boolean): SlideElement[] => {
        const spTree = first(first(part?.root, "cSld"), "spTree");
        return spTree
          ? parseTree(
              spTree,
              { ...context, part: part! },
              {
                template,
                inherit,
                defaultText: context.defaultText,
              },
            )
          : [];
      };
      const showTemplates = attr(slide.root, "showMasterSp") !== "0";
      const showMaster = showTemplates && attr(layout?.root, "showMasterSp") !== "0";
      return {
        number: index + 1,
        hidden: attr(slide.root, "show") === "0",
        background: background(
          [master, layout, slide].filter((part): part is Part => !!part),
          context,
        ),
        elements: [
          ...(showMaster ? tree(master, true) : []),
          ...(showTemplates ? tree(layout, true) : []),
          ...tree(slide, false),
        ],
      };
    }),
  );
  const media = await readMedia(
    opc,
    slides.flatMap((slide) => [
      ...(slide.background?.type === "image" ? [slide.background.part] : []),
      ...mediaParts(slide.elements),
    ]),
  );
  return {
    width: emu(attr(size, "cx")) ?? 720,
    height: emu(attr(size, "cy")) ?? 540,
    slides,
    media,
  };
};

const readPart = async (opc: OpcPackage, path: string): Promise<Part> => {
  const [doc, rels] = await Promise.all([opc.xml(path), opc.relationships(path)]);
  return { opc, path, rels, root: doc.documentElement };
};
