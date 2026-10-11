import type { Block, DocxDocument } from "../docx/model";
import type { RenderContext } from "./context";
import { addComments } from "./comments";
import { renderParagraph } from "./paragraph";
import { renderTable } from "./table";
import { parseDocx } from "../docx/parse";
import { paginate } from "./paginate";
import { resolveDocumentParagraphMarks } from "./revisions";

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

interface RenderLease {
  cancel(): void;
}

const renderLeases = new WeakMap<HTMLElement, RenderLease>();

const acquireContainerLease = (container: HTMLElement, onCancel: () => void): RenderLease => {
  let cancelled = false;
  const lease: RenderLease = {
    cancel() {
      if (cancelled) return;
      cancelled = true;
      onCancel();
      if (renderLeases.get(container) === lease) renderLeases.delete(container);
    },
  };
  const previous = renderLeases.get(container);
  renderLeases.set(container, lease);
  previous?.cancel();
  return lease;
};

const abortError = (): DOMException => {
  return new DOMException("Rendering was aborted.", "AbortError");
};

export const renderDocument = (
  doc: DocxDocument,
  container: HTMLElement,
  options: RenderDocumentOptions = {},
): RenderedDocument => {
  const objectUrls = new Set<string>();
  const objectUrlsByPart = new Map<string, string>();
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
    objectUrlsByPart,
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
    for (const url of objectUrls) URL.revokeObjectURL(url);
    objectUrls.clear();
    objectUrlsByPart.clear();
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
        if (cancelled || renderLeases.get(container) !== lease) return;
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
  input: ArrayBuffer | Uint8Array | Blob,
  container: HTMLElement,
  options: RenderOptions = {},
): Promise<RenderedDocument> => {
  const controller = new AbortController();
  let cancelled = false;
  let lease: RenderLease;
  const onAbort = () => lease.cancel();
  lease = acquireContainerLease(container, () => {
    cancelled = true;
    controller.abort();
    options.signal?.removeEventListener("abort", onAbort);
  });
  if (options.signal?.aborted) lease.cancel();
  else options.signal?.addEventListener("abort", onAbort, { once: true });

  const check = () => {
    if (cancelled || controller.signal.aborted || renderLeases.get(container) !== lease)
      throw abortError();
  };

  let doc: DocxDocument;
  try {
    check();
    doc = await parseDocx(input);
    check();
  } catch (error) {
    const aborted = cancelled || controller.signal.aborted || renderLeases.get(container) !== lease;
    lease.cancel();
    if (aborted) throw abortError();
    throw error;
  }

  return renderDocument(doc, container, {
    changes: options.changes,
    comments: options.comments,
  });
};
