import { strToU8, zipSync } from "fflate";

export const PML = "http://schemas.openxmlformats.org/presentationml/2006/main";
export const DML = "http://schemas.openxmlformats.org/drawingml/2006/main";
export const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELS = "http://schemas.openxmlformats.org/package/2006/relationships";
const namespaces = `xmlns:p="${PML}" xmlns:a="${DML}" xmlns:r="${REL}"`;

export interface PptxRelationship {
  readonly id: string;
  /** Relationship type suffix (e.g. `image`) or a full type URI. */
  readonly type: string;
  readonly target: string;
  readonly external?: boolean;
}

export interface PptxSlidePart {
  /** Children of `p:spTree`. */
  readonly shapes?: string;
  /** Extra attributes on the root element (e.g. `show="0"`). */
  readonly attributes?: string;
  /** Children of `p:cSld` placed before `p:spTree` (e.g. `p:bg`). */
  readonly background?: string;
  /** Children placed after `p:cSld` (e.g. `p:clrMapOvr`, `p:txStyles`). */
  readonly extra?: string;
  readonly relationships?: readonly PptxRelationship[];
}

export interface BuildPptxOptions {
  readonly presentationPath?: string;
  readonly size?: { readonly cx: number; readonly cy: number };
  readonly slides?: readonly PptxSlidePart[];
  readonly master?: PptxSlidePart;
  readonly layout?: PptxSlidePart;
  /** Full theme XML; a default Office-like theme is used when omitted. */
  readonly theme?: string;
  readonly tableStyles?: string;
  readonly defaultTextStyle?: string;
  readonly media?: Readonly<Record<string, Uint8Array>>;
}

export const defaultTheme = `<a:theme xmlns:a="${DML}" name="Test"><a:themeElements>
<a:clrScheme name="Test"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
<a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4>
<a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink>
<a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme>
<a:fontScheme name="Test"><a:majorFont><a:latin typeface="Major Sans"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
<a:minorFont><a:latin typeface="Minor Sans"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>
<a:fmtScheme name="Test"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:shade val="50000"/></a:schemeClr></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
<a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/></a:schemeClr></a:solidFill></a:ln><a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="38100"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="50000"/></a:schemeClr></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
</a:fmtScheme></a:themeElements></a:theme>`;

const defaultClrMap = `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>`;

/** A `p:sp` with an optional placeholder, transform and text body. */
export const shape = (options: {
  readonly id: number;
  readonly name?: string;
  readonly placeholder?: string;
  readonly xfrm?: readonly [number, number, number, number];
  readonly properties?: string;
  readonly body?: string;
  readonly style?: string;
  readonly nonVisual?: string;
}): string =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${options.id}" name="${options.name ?? `Shape ${options.id}`}"${options.nonVisual ?? ""}/><p:cNvSpPr/><p:nvPr>${options.placeholder ?? ""}</p:nvPr></p:nvSpPr>` +
  `<p:spPr>${options.xfrm ? `<a:xfrm><a:off x="${options.xfrm[0]}" y="${options.xfrm[1]}"/><a:ext cx="${options.xfrm[2]}" cy="${options.xfrm[3]}"/></a:xfrm>` : ""}${options.properties ?? ""}</p:spPr>` +
  `${options.style ?? ""}${options.body === undefined ? "" : `<p:txBody><a:bodyPr/><a:lstStyle/>${options.body}</p:txBody>`}</p:sp>`;

const relsPath = (path: string): string => {
  const parts = path.split("/");
  const file = parts.pop() ?? "";
  return [...parts, "_rels", `${file}.rels`].join("/");
};

const relationships = (items: readonly PptxRelationship[]): string =>
  `<Relationships xmlns="${RELS}">${items
    .map(
      (item) =>
        `<Relationship Id="${item.id}" Type="${item.type.includes("/") ? item.type : `${REL}/${item.type}`}" Target="${item.target}"${item.external ? ' TargetMode="External"' : ""}/>`,
    )
    .join("")}</Relationships>`;

const part = (
  root: "sld" | "sldLayout" | "sldMaster",
  input: PptxSlidePart = {},
  fallback = "",
): string =>
  `<p:${root} ${namespaces}${input.attributes ? ` ${input.attributes}` : ""}><p:cSld>${input.background ?? ""}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${input.shapes ?? ""}</p:spTree></p:cSld>${input.extra ?? fallback}</p:${root}>`;

/** Builds a minimal PPTX package with one master, one layout and the given slides. */
export const buildPptx = (options: BuildPptxOptions = {}): Uint8Array => {
  const presentationPath = options.presentationPath ?? "ppt/presentation.xml";
  const base = presentationPath.split("/").slice(0, -1).join("/");
  const at = (path: string) => (base ? `${base}/${path}` : path);
  const archive: Record<string, Uint8Array> = {};
  const types: string[] = [];
  const add = (path: string, xml: string, type?: string, rels?: readonly PptxRelationship[]) => {
    archive[path] = strToU8(xml);
    if (type)
      types.push(
        `<Override PartName="/${path}" ContentType="application/vnd.openxmlformats-officedocument.${type}"/>`,
      );
    if (rels?.length) archive[relsPath(path)] = strToU8(relationships(rels));
  };
  const slides = options.slides ?? [];
  add(at("theme/theme1.xml"), options.theme ?? defaultTheme, "theme+xml");
  add(
    at("slideMasters/slideMaster1.xml"),
    part("sldMaster", options.master, defaultClrMap),
    "presentationml.slideMaster+xml",
    [
      { id: "rLayout", type: "slideLayout", target: "../slideLayouts/slideLayout1.xml" },
      { id: "rTheme", type: "theme", target: "../theme/theme1.xml" },
      ...(options.master?.relationships ?? []),
    ],
  );
  add(
    at("slideLayouts/slideLayout1.xml"),
    part("sldLayout", options.layout),
    "presentationml.slideLayout+xml",
    [
      { id: "rMaster", type: "slideMaster", target: "../slideMasters/slideMaster1.xml" },
      ...(options.layout?.relationships ?? []),
    ],
  );
  slides.forEach((slide, index) =>
    add(at(`slides/slide${index + 1}.xml`), part("sld", slide), "presentationml.slide+xml", [
      { id: "rLayout", type: "slideLayout", target: "../slideLayouts/slideLayout1.xml" },
      ...(slide.relationships ?? []),
    ]),
  );
  for (const [path, bytes] of Object.entries(options.media ?? {})) archive[at(path)] = bytes;
  if (options.tableStyles)
    add(at("tableStyles.xml"), options.tableStyles, "presentationml.tableStyles+xml");
  const size = options.size ?? { cx: 9144000, cy: 6858000 };
  add(
    presentationPath,
    `<p:presentation ${namespaces}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rMaster"/></p:sldMasterIdLst><p:sldIdLst>${slides
      // happy-dom keeps only the first of two attributes sharing a local name, so r:id goes first.
      .map((_, index) => `<p:sldId r:id="rSlide${index + 1}" id="${256 + index}"/>`)
      .join(
        "",
      )}</p:sldIdLst><p:sldSz cx="${size.cx}" cy="${size.cy}"/>${options.defaultTextStyle ?? ""}</p:presentation>`,
    "presentationml.presentation.main+xml",
    [
      { id: "rMaster", type: "slideMaster", target: "slideMasters/slideMaster1.xml" },
      { id: "rTheme", type: "theme", target: "theme/theme1.xml" },
      ...slides.map((_, index) => ({
        id: `rSlide${index + 1}`,
        type: "slide",
        target: `slides/slide${index + 1}.xml`,
      })),
      ...(options.tableStyles
        ? [{ id: "rTableStyles", type: "tableStyles", target: "tableStyles.xml" }]
        : []),
    ],
  );
  archive["_rels/.rels"] = strToU8(
    relationships([{ id: "rDoc", type: "officeDocument", target: presentationPath }]),
  );
  archive["[Content_Types].xml"] = strToU8(
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>${types.join("")}</Types>`,
  );
  return zipSync(archive, { level: 0 });
};
