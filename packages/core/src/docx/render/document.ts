import type { Block, DocxDocument } from "../model";
import type { RenderContext } from "./context";
import { addComments } from "./comments";
import { renderParagraph } from "./paragraph";
import { renderTable } from "./table";
import { parseDocx } from "../parse";
import { paginate } from "./paginate";
import { resolveDocumentParagraphMarks } from "./revisions";
import { acquireContainerLease } from "../../common/lease";
import { createObjectUrls } from "../../common/media";
import { renderParsed } from "../../common/render-file";
import type { PackageInput } from "../../opc/package";

export type ChangesMode = "markup" | "final" | "original";

export interface RenderOptions {
  changes?: ChangesMode;
  comments?: boolean;
  signal?: AbortSignal;
}

export interface RenderDocumentOptions {
  changes?: ChangesMode;
  comments?: boolean;
}

export interface RenderedDocument {
  readonly document: DocxDocument;
  update(options: { changes?: ChangesMode; comments?: boolean }): void;
  destroy(): void;
}

const renderBlock = (
  block: Block,
  context: RenderContext,
  pageNumber?: number,
  decoration = false,
): HTMLElement => {
  const scopedContext = decoration ? { ...context, decorationPageNumber: pageNumber } : context;
  if (block.type === "paragraph") return renderParagraph(block, scopedContext);
  return renderTable(block, scopedContext, (nested, nestedContext) =>
    renderBlock(nested, nestedContext, pageNumber, decoration),
  );
};

export const renderDocument = (
  doc: DocxDocument,
  container: HTMLElement,
  options: RenderDocumentOptions = {},
): RenderedDocument => {
  const objectUrls = createObjectUrls();
  const fields: RenderContext["fields"] = [];
  const commentAnchors: RenderContext["commentAnchors"] = new Map();
  const commentBalloons: RenderContext["commentBalloons"] = new Map();
  const commentReferences: RenderContext["commentReferences"] = new Map();
  let cancelled = false;
  let root: HTMLDivElement | undefined;
  let removeCommentListeners = () => {};
  let context: RenderContext = {
    doc,
    comments: options.comments !== false,
    changes: options.changes ?? "markup",
    objectUrls,
    colors: new Map(),
    fields,
    commentAnchors,
    commentBalloons,
    commentReferences,
  };
  const lease = acquireContainerLease(container, () => {
    cancelled = true;
    removeCommentListeners();
    if (root?.parentElement === container) container.replaceChildren();
    objectUrls.revoke();
  });

  try {
    root = document.createElement("div");
    if (doc.defaultTabStop !== undefined)
      root.style.setProperty("--orim-tab-size", `${doc.defaultTabStop}pt`);
    fields.length = 0;
    commentAnchors.clear();
    commentBalloons.clear();
    commentReferences.clear();
    const draw = () => {
      removeCommentListeners();
      fields.length = 0;
      commentAnchors.clear();
      commentBalloons.clear();
      commentReferences.clear();
      root!.className = `orim-root orim-${context.changes}`;
      const resolvedDoc = resolveDocumentParagraphMarks(doc, context.changes);
      const pages = paginate(
        root!,
        resolvedDoc,
        context.changes,
        (block, pageNumber, decoration) => renderBlock(block, context, pageNumber, decoration),
        context.comments,
        fields,
      );
      if (context.comments)
        removeCommentListeners = addComments(
          pages,
          doc.comments,
          commentAnchors,
          commentBalloons,
          root!,
        );
    };
    container.replaceChildren(root);
    draw();
    return {
      document: doc,
      destroy() {
        lease.cancel();
      },
      update(next) {
        if (cancelled || !lease.active) return;
        const changes = next.changes ?? context.changes;
        const comments = next.comments ?? context.comments;
        if (changes === context.changes && comments === context.comments) return;
        context = { ...context, changes, comments };
        draw();
      },
    };
  } catch (error) {
    lease.cancel();
    throw error;
  }
};

export const renderDocx = async (
  input: PackageInput,
  container: HTMLElement,
  options: RenderOptions = {},
): Promise<RenderedDocument> => {
  return renderParsed(
    () => parseDocx(input),
    container,
    options.signal,
    (doc) =>
      renderDocument(doc, container, { changes: options.changes, comments: options.comments }),
  );
};
