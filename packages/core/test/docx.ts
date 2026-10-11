import { strToU8, zipSync } from "fflate";

export interface BuildDocxOptions {
  document: string;
  mainPath?: string;
  styles?: string;
  stylesPath?: string;
  numbering?: string;
  numberingPath?: string;
  header?: string;
  headerPath?: string;
  comments?: string;
  commentsPath?: string;
  commentsExtended?: string;
  commentsExtendedPath?: string;
  theme?: string;
  themePath?: string;
  relationships?: string;
  media?: Record<string, { bytes: Uint8Array; mime: string }>;
}

export const buildDocx = (options: BuildDocxOptions): Uint8Array => {
  const mainPath = options.mainPath ?? "word/document.xml";
  const mainDir = mainPath.slice(0, mainPath.lastIndexOf("/"));
  const relativeTarget = (path: string) =>
    mainDir && path.startsWith(`${mainDir}/`) ? path.slice(mainDir.length + 1) : path;
  const parts: Record<string, string> = {
    [mainPath]: options.document,
  };
  const stylesPath = options.stylesPath ?? "word/styles.xml";
  const numberingPath = options.numberingPath ?? "word/numbering.xml";
  const headerPath = options.headerPath ?? "word/header1.xml";
  const commentsPath = options.commentsPath ?? "word/comments.xml";
  const commentsExtendedPath = options.commentsExtendedPath ?? "word/commentsExtended.xml";
  const themePath = options.themePath ?? "word/theme/theme1.xml";
  if (options.styles) parts[stylesPath] = options.styles;
  if (options.numbering) parts[numberingPath] = options.numbering;
  if (options.header) parts[headerPath] = options.header;
  if (options.comments) parts[commentsPath] = options.comments;
  if (options.commentsExtended) parts[commentsExtendedPath] = options.commentsExtended;
  if (options.theme) parts[themePath] = options.theme;
  const overrides = Object.keys(parts)
    .filter((path) => path !== mainPath)
    .map(
      (path) =>
        `<Override PartName="/${path}" ContentType="${path === commentsExtendedPath ? "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml" : path === headerPath ? "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml" : "application/xml"}"/>`,
    )
    .join("");
  const media = Object.entries(options.media ?? {});
  const mediaTypes = new Map(
    media.map(([path, part]) => [path.slice(path.lastIndexOf(".") + 1), part.mime]),
  );
  const mediaDefaults = [...mediaTypes]
    .map(([extension, mime]) => `<Default Extension="${extension}" ContentType="${mime}"/>`)
    .join("");
  const relationships = [
    options.styles &&
      `<Relationship Id="rStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="${relativeTarget(stylesPath)}"/>`,
    options.numbering &&
      `<Relationship Id="rNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="${relativeTarget(numberingPath)}"/>`,
    options.header &&
      `<Relationship Id="rHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="${relativeTarget(headerPath)}"/>`,
    options.comments &&
      `<Relationship Id="rComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="${relativeTarget(commentsPath)}"/>`,
    options.commentsExtended &&
      `<Relationship Id="rCommentsExtended" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="${relativeTarget(commentsExtendedPath)}"/>`,
    options.theme &&
      `<Relationship Id="rTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="${relativeTarget(themePath)}"/>`,
  ]
    .filter(Boolean)
    .concat(options.relationships ?? "")
    .join("");
  const archive: Record<string, Uint8Array> = {
    "[Content_Types].xml": strToU8(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${mediaDefaults}<Override PartName="/${mainPath}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>${overrides}</Types>`,
    ),
    "_rels/.rels": strToU8(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="${mainPath}"/></Relationships>`,
    ),
    [`${mainDir}/_rels/${mainPath.slice(mainDir.length + 1)}.rels`]: strToU8(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships}</Relationships>`,
    ),
  };
  for (const [path, xml] of Object.entries(parts)) archive[path] = strToU8(xml);
  for (const [path, part] of media) archive[path] = part.bytes;
  return zipSync(archive, { level: 0 });
};
