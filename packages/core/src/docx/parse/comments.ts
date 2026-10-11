import type { Block, DocxComment } from "../model";
import { attr, children, descendants, elements, W14, W15 } from "../../xml";
import type { OpcPackage } from "../../opc/package";
import type { ParseContext } from "./context";
import { createParseState } from "./context";

type ParseBlocks = (nodes: Element[], context: ParseContext) => Block[];

export async function parseComments(
  opc: OpcPackage,
  context: ParseContext,
  parseBlocks: ParseBlocks,
): Promise<DocxComment[]> {
  const config = context.config;
  const relatedPath = (suffix: string): string | undefined => {
    const relation = [...config.rels.values()].find(
      (item) => !item.external && item.type.endsWith(`/${suffix}`),
    );
    return relation ? opc.resolve(config.source, relation.target) : undefined;
  };
  const commentsExtendedPath = relatedPath("commentsExtended");
  const extended =
    commentsExtendedPath && opc.zip.has(commentsExtendedPath)
      ? await opc.xml(commentsExtendedPath)
      : undefined;
  const extensionByParaId = new Map(
    (extended ? descendants(extended, "commentEx") : []).flatMap((node) => {
      const paraId = attr(node, "paraId", W15);
      return paraId
        ? [
            [
              paraId,
              {
                parentParaId: attr(node, "paraIdParent", W15),
                done: attr(node, "done", W15) === "1" || attr(node, "done", W15) === "true",
              },
            ] as const,
          ]
        : [];
    }),
  );
  const commentsPath = relatedPath("comments");
  if (!commentsPath || !opc.zip.has(commentsPath)) return [];

  const commentsDoc = await opc.xml(commentsPath);
  const nodes = descendants(commentsDoc, "comment");
  const nodesByParaId = new Map(
    nodes.flatMap((node) => {
      const paraId = elements(node, "p").at(-1) && attr(elements(node, "p").at(-1), "paraId", W14);
      return paraId ? [[paraId, node] as const] : [];
    }),
  );
  const commentContext: ParseContext = {
    config: {
      ...config,
      source: commentsPath,
      rels: await opc.relationships(commentsPath),
    },
    state: createParseState(),
  };
  return nodes.flatMap((node) => {
    const id = attr(node, "id");
    if (!id) return [];
    const lastParagraph = elements(node, "p").at(-1);
    const paraId = lastParagraph && attr(lastParagraph, "paraId", W14);
    const extension = paraId ? extensionByParaId.get(paraId) : undefined;
    const parentNode = extension?.parentParaId
      ? nodesByParaId.get(extension.parentParaId)
      : undefined;
    return [
      {
        id,
        author: attr(node, "author") ?? "Unknown",
        initials: attr(node, "initials"),
        date: attr(node, "date"),
        paragraphs: parseBlocks(children(node), {
          ...commentContext,
          state: createParseState(),
        }),
        paraId,
        ...(extension
          ? { parentId: parentNode ? attr(parentNode, "id") : undefined, done: extension.done }
          : {}),
      },
    ];
  });
}
