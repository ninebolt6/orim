import { attr, descendants, R } from "../common/xml";
import type { OpcPackage, Relationship } from "../opc/package";

/** Resolves the package part embedded by the first `a:blip` under `element`. */
export const blipPart = (
  element: Element,
  rels: ReadonlyMap<string, Relationship>,
  source: string,
  opc: OpcPackage,
): string | undefined => {
  const blip = element.localName === "blip" ? element : descendants(element, "blip")[0];
  const id = blip && (attr(blip, "embed", R) ?? attr(blip, "link", R));
  const relation = id ? rels.get(id) : undefined;
  return relation && !relation.external ? opc.resolve(source, relation.target) : undefined;
};
