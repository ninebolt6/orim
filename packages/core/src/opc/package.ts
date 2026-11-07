import { openZip, type ZipArchive } from "./zip";
import { descendants } from "../common/xml";
import type { MediaPart } from "../common/media";

export type PackageInput = ArrayBuffer | Uint8Array | Blob;

export interface Relationship {
  id: string;
  type: string;
  target: string;
  external: boolean;
}

export interface OpcPackage {
  zip: ZipArchive;
  xml(path: string): Promise<Document>;
  bytes(path: string): Promise<Uint8Array>;
  relationships(path: string): Promise<Map<string, Relationship>>;
  resolve(source: string, target: string): string;
  mime(path: string): Promise<string>;
}

export const findRelationship = (
  relationships: ReadonlyMap<string, Relationship>,
  type: string,
): Relationship | undefined => {
  for (const relationship of relationships.values())
    if (!relationship.external && relationship.type.endsWith(`/${type}`)) return relationship;
  return undefined;
};

const normalize = (path: string): string => {
  const parts: string[] = [];
  for (const part of path.split("/")) {
    if (!part || part === ".") continue;
    if (part === "..") parts.pop();
    else parts.push(part);
  }
  return parts.join("/");
};

export const openOpc = (input: ArrayBuffer | Uint8Array): OpcPackage => {
  const zip = openZip(input);
  const relCache = new Map<string, Promise<Map<string, Relationship>>>();
  const contentTypes = new Map<string, string>();
  const types = zip.has("[Content_Types].xml")
    ? zip.text("[Content_Types].xml").then((xml) => {
        const doc = new DOMParser().parseFromString(xml, "application/xml");
        for (const node of descendants(doc, "Default")) {
          const ext = node.getAttribute("Extension");
          const type = node.getAttribute("ContentType");
          if (ext && type) contentTypes.set(`.${ext.toLowerCase()}`, type);
        }
        for (const node of descendants(doc, "Override")) {
          const part = node.getAttribute("PartName");
          const type = node.getAttribute("ContentType");
          if (part && type) contentTypes.set(normalize(part), type);
        }
      })
    : Promise.resolve();

  const relPath = (source: string): string => {
    const normalized = normalize(source);
    if (!normalized) return "_rels/.rels";
    const path = normalized.split("/");
    const file = path.pop() ?? "";
    return [...path, "_rels", `${file}.rels`].join("/");
  };

  return {
    zip,
    async bytes(path) {
      return zip.read(normalize(path));
    },
    async xml(path) {
      const text = await zip.text(normalize(path));
      const doc = new DOMParser().parseFromString(text, "application/xml");
      if (doc.querySelector("parsererror")) throw new Error(`Invalid XML part: ${path}`);
      return doc;
    },
    resolve(source, target) {
      if (target.startsWith("/")) return normalize(target);
      const base = source.split("/").slice(0, -1).join("/");
      return normalize(`${base}/${target}`);
    },
    relationships(source) {
      let pending = relCache.get(source);
      if (!pending) {
        pending = (async () => {
          const out = new Map<string, Relationship>();
          const path = relPath(source);
          if (!zip.has(path)) return out;
          const doc = await this.xml(path);
          for (const rel of descendants(doc, "Relationship")) {
            const id = rel.getAttribute("Id");
            const type = rel.getAttribute("Type");
            const target = rel.getAttribute("Target");
            if (id && type && target) {
              out.set(id, {
                id,
                type,
                target,
                external: rel.getAttribute("TargetMode") === "External",
              });
            }
          }
          return out;
        })();
        relCache.set(source, pending);
      }
      return pending;
    },
    async mime(path) {
      await types;
      return (
        contentTypes.get(normalize(path)) ??
        contentTypes.get(path.slice(path.lastIndexOf(".")).toLowerCase()) ??
        "application/octet-stream"
      );
    },
  };
};

export const loadOpc = async (input: PackageInput): Promise<OpcPackage> => {
  return openOpc(
    typeof Blob !== "undefined" && input instanceof Blob
      ? await input.arrayBuffer()
      : (input as ArrayBuffer | Uint8Array),
  );
};

export const readMedia = async (
  opc: OpcPackage,
  paths: Iterable<string>,
): Promise<Readonly<Record<string, MediaPart>>> => {
  return Object.fromEntries(
    await Promise.all(
      [...new Set(paths)].map(async (path) => {
        const [bytes, mime] = await Promise.all([opc.bytes(path), opc.mime(path)]);
        return [path, { bytes, mime }] as const;
      }),
    ),
  );
};

/** Resolves the first internal part related to `source` by relationship `type`. */
export const relatedPart = async (
  opc: OpcPackage,
  source: string,
  type: string,
): Promise<string | undefined> => {
  const relationship = findRelationship(await opc.relationships(source), type);
  return relationship && opc.resolve(source, relationship.target);
};

export const optionalXml = async (
  opc: OpcPackage,
  path: string | undefined,
): Promise<Document | undefined> => {
  return path && opc.zip.has(path) ? opc.xml(path) : undefined;
};
