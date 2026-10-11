import type { Block, DocxComment } from "../model";
import { attr, children, descendants, elements } from "../../xml";
import type { OpcPackage } from "../../opc/package";
import type { ParseContext } from "./context";
import { createFieldStack } from "./fields";

type ParseBlocks = (nodes: Element[], context: ParseContext, comments: Set<string>) => Block[];

export async function parseComments(
  opc: OpcPackage,
  context: ParseContext,
  parseBlocks: ParseBlocks,
): Promise<DocxComment[]> {
  const commentList: DocxComment[] = [];
  if (opc.zip.has("word/comments.xml")) {
    const commentsPath = "word/comments.xml";
    const commentsDoc = await opc.xml(commentsPath);
    const commentContext: ParseContext = {
      ...context,
      source: commentsPath,
      rels: await opc.relationships(commentsPath),
      commentsOpen: new Set(),
      fields: createFieldStack(),
    };
    for (const node of descendants(commentsDoc, "comment")) {
      const id = attr(node, "id");
      if (!id) continue;
      const blocks = parseBlocks(children(node), commentContext, new Set());
      const lastParagraph = elements(node, "p").at(-1);
      const paraId = lastParagraph && attr(lastParagraph, "paraId");
      commentList.push({
        id,
        author: attr(node, "author") ?? "Unknown",
        initials: attr(node, "initials"),
        date: attr(node, "date"),
        paragraphs: blocks,
        paraId,
      });
    }
  }
  if (opc.zip.has("word/commentsExtended.xml")) {
    const extended = await opc.xml("word/commentsExtended.xml");
    const byPara = new Map(commentList.map((comment) => [comment.paraId ?? "", comment]));
    for (const node of descendants(extended, "commentEx")) {
      const paraId = attr(node, "paraId");
      const parent = attr(node, "paraIdParent");
      const comment = byPara.get(paraId ?? "");
      if (!comment) continue;
      comment.parentId = byPara.get(parent ?? "")?.id;
      comment.done = attr(node, "done") === "1" || attr(node, "done") === "true";
    }
  }
  return commentList;
}
