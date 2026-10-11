import type { Block, DocxComment } from "../model";
import { attr, children, descendants, elements } from "../../xml";
import type { OpcPackage } from "../../opc/package";
import type { ParseContext } from "./context";
import { createCommentRangeState } from "./context";
import { createFieldStack } from "./fields";

type ParseBlocks = (nodes: Element[], context: ParseContext) => Block[];

export async function parseComments(
  opc: OpcPackage,
  context: ParseContext,
  parseBlocks: ParseBlocks,
): Promise<DocxComment[]> {
  const extended = opc.zip.has("word/commentsExtended.xml")
    ? await opc.xml("word/commentsExtended.xml")
    : undefined;
  const extensionByParaId = new Map(
    (extended ? descendants(extended, "commentEx") : []).flatMap((node) => {
      const paraId = attr(node, "paraId");
      return paraId
        ? [
            [
              paraId,
              {
                parentParaId: attr(node, "paraIdParent"),
                done: attr(node, "done") === "1" || attr(node, "done") === "true",
              },
            ] as const,
          ]
        : [];
    }),
  );
  if (!opc.zip.has("word/comments.xml")) return [];

  const commentsPath = "word/comments.xml";
  const commentsDoc = await opc.xml(commentsPath);
  const nodes = descendants(commentsDoc, "comment");
  const nodesByParaId = new Map(
    nodes.flatMap((node) => {
      const paraId = elements(node, "p").at(-1) && attr(elements(node, "p").at(-1), "paraId");
      return paraId ? [[paraId, node] as const] : [];
    }),
  );
  const commentContext: ParseContext = {
    ...context,
    source: commentsPath,
    rels: await opc.relationships(commentsPath),
    commentRanges: createCommentRangeState(),
    fields: createFieldStack(),
  };
  return nodes.flatMap((node) => {
    const id = attr(node, "id");
    if (!id) return [];
    const lastParagraph = elements(node, "p").at(-1);
    const paraId = lastParagraph && attr(lastParagraph, "paraId");
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
          commentRanges: createCommentRangeState(),
        }),
        paraId,
        ...(extension
          ? { parentId: parentNode ? attr(parentNode, "id") : undefined, done: extension.done }
          : {}),
      },
    ];
  });
}
