import { strToU8, zipSync } from "fflate";

export interface BuildDocxOptions {
  document: string;
  styles?: string;
  numbering?: string;
  comments?: string;
  commentsExtended?: string;
  relationships?: string;
  media?: Record<string, { bytes: Uint8Array; mime: string }>;
}

export function buildDocx(options: BuildDocxOptions): Uint8Array {
  const parts: Record<string, string> = {
    "word/document.xml": options.document,
  };
  if (options.styles) parts["word/styles.xml"] = options.styles;
  if (options.numbering) parts["word/numbering.xml"] = options.numbering;
  if (options.comments) parts["word/comments.xml"] = options.comments;
  if (options.commentsExtended) parts["word/commentsExtended.xml"] = options.commentsExtended;
  const overrides = Object.keys(parts)
    .filter((path) => path !== "word/document.xml")
    .map((path) => `<Override PartName="/${path}" ContentType="application/xml"/>`)
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
      '<Relationship Id="rStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
    options.numbering &&
      '<Relationship Id="rNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>',
    options.comments &&
      '<Relationship Id="rComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>',
    options.commentsExtended &&
      '<Relationship Id="rCommentsExtended" Type="http://schemas.microsoft.com/office/word/2018/wordml/commentsExtended" Target="commentsExtended.xml"/>',
  ]
    .filter(Boolean)
    .concat(options.relationships ?? "")
    .join("");
  const archive: Record<string, Uint8Array> = {
    "[Content_Types].xml": strToU8(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${mediaDefaults}<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>${overrides}</Types>`,
    ),
    "_rels/.rels": strToU8(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    ),
    "word/_rels/document.xml.rels": strToU8(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships}</Relationships>`,
    ),
  };
  for (const [path, xml] of Object.entries(parts)) archive[path] = strToU8(xml);
  for (const [path, part] of media) archive[path] = part.bytes;
  return zipSync(archive, { level: 0 });
}
