import { openZip, type ZipArchive } from "./zip";

export interface Relationship {
  id: string;
  type: string;
  target: string;
  external: boolean;
}

export interface ThemeFonts {
  major?: string;
  minor?: string;
  majorEastAsia?: string;
  minorEastAsia?: string;
}

export interface OpcPackage {
  zip: ZipArchive;
  xml(path: string): Promise<Document>;
  bytes(path: string): Promise<Uint8Array>;
  relationships(path: string): Promise<Map<string, Relationship>>;
  resolve(source: string, target: string): string;
  themeFonts(): Promise<ThemeFonts>;
  mime(path: string): Promise<string>;
}

export function elements(parent: ParentNode, name: string): Element[] {
  return Array.from(parent.childNodes).filter(
    (node): node is Element => node.nodeType === 1 && (node as Element).localName === name,
  );
}

export function descendants(parent: ParentNode, name: string): Element[] {
  const result: Element[] = [];
  const visit = (node: ParentNode): void => {
    for (const child of Array.from(node.childNodes)) {
      if (child.nodeType !== 1) continue;
      const element = child as Element;
      if (element.localName === name) result.push(element);
      visit(element);
    }
  };
  visit(parent);
  return result;
}

export function first(parent: ParentNode, name: string): Element | undefined {
  return elements(parent, name)[0];
}

export function attr(element: Element | undefined, name: string): string | undefined {
  if (!element) return undefined;
  return element.getAttribute(`w:${name}`) ?? element.getAttribute(name) ?? undefined;
}

function normalize(path: string): string {
  const parts: string[] = [];
  for (const part of path.split("/")) {
    if (!part || part === ".") continue;
    if (part === "..") parts.pop();
    else parts.push(part);
  }
  return parts.join("/");
}

export function openOpc(input: ArrayBuffer | Uint8Array): OpcPackage {
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

  function relPath(source: string): string {
    const path = source.split("/");
    const file = path.pop() ?? "";
    return [...path, "_rels", `${file}.rels`].join("/");
  }

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
    async themeFonts() {
      for (const path of zip.paths()) {
        if (!path.startsWith("word/theme/") || !path.endsWith(".xml")) continue;
        const doc = await this.xml(path);
        const major = descendants(doc, "majorFont")[0];
        const minor = descendants(doc, "minorFont")[0];
        const typeface = (font: Element | undefined, name: string) =>
          font ? (elements(font, name)[0]?.getAttribute("typeface") ?? undefined) : undefined;
        return {
          major: typeface(major, "latin"),
          minor: typeface(minor, "latin"),
          majorEastAsia: typeface(major, "ea"),
          minorEastAsia: typeface(minor, "ea"),
        };
      }
      return {};
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
}
