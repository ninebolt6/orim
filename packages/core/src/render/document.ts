import css from "./styles.css?inline";
import type { Block, DocxDocument } from "../docx/model";
import type { RenderContext } from "./context";
import { addComments } from "./comments";
import { renderParagraph } from "./paragraph";
import { renderTable } from "./table";
import { parseDocx } from "../docx/parse";
import { paginate } from "./paginate";

export type ChangesMode = "markup" | "final" | "original";

export interface RenderOptions {
  changes?: ChangesMode;
  comments?: boolean;
  signal?: AbortSignal;
}

export interface RenderedDocument {
  readonly document: DocxDocument;
  update(options: { changes?: ChangesMode; comments?: boolean }): void;
  destroy(): void;
}

function renderBlock(block: Block, context: RenderContext): HTMLElement {
  if (block.type === "paragraph") return renderParagraph(block, context);
  return renderTable(block, context, renderBlock);
}

function ensureStyles(): void {
  if (document.querySelector("style[data-orim]")) return;
  const style = document.createElement("style");
  style.dataset.orim = "true";
  style.textContent = css;
  document.head.append(style);
}

function isModel(input: unknown): input is DocxDocument {
  return !!input && typeof input === "object" && Array.isArray((input as DocxDocument).sections);
}

interface RenderLease {
  cancel(): void;
}

const renderLeases = new WeakMap<HTMLElement, RenderLease>();

function abortError(): DOMException {
  return new DOMException("Rendering was aborted.", "AbortError");
}

export async function renderDocx(
  input: ArrayBuffer | Uint8Array | Blob | DocxDocument,
  container: HTMLElement,
  options: RenderOptions = {},
): Promise<RenderedDocument> {
  ensureStyles();
  const objectUrls = new Set<string>();
  const controller = new AbortController();
  let cancelled = false;
  let root: HTMLDivElement | undefined;
  const onAbort = () => lease.cancel();
  const lease: RenderLease = {
    cancel() {
      if (cancelled) return;
      cancelled = true;
      controller.abort();
      options.signal?.removeEventListener("abort", onAbort);
      if (root?.parentElement === container) container.replaceChildren();
      for (const url of objectUrls) URL.revokeObjectURL(url);
      objectUrls.clear();
      if (renderLeases.get(container) === lease) renderLeases.delete(container);
    },
  };
  const previous = renderLeases.get(container);
  renderLeases.set(container, lease);
  previous?.cancel();
  if (options.signal?.aborted) lease.cancel();
  else options.signal?.addEventListener("abort", onAbort, { once: true });

  const check = () => {
    if (cancelled || controller.signal.aborted || renderLeases.get(container) !== lease)
      throw abortError();
  };

  try {
    check();
    const doc = isModel(input) ? input : await parseDocx(input);
    check();
    root = document.createElement("div");
    const context: RenderContext = {
      doc,
      comments: options.comments !== false,
      changes: options.changes ?? "markup",
      objectUrls,
      colors: new Map(),
    };
    if (doc.defaultTabStop !== undefined)
      root.style.setProperty("--orim-tab-size", `${doc.defaultTabStop}pt`);
    const draw = () => {
      root!.className = `orim-root orim-${context.changes}`;
      const pages = paginate(
        root!,
        doc,
        context.changes,
        (block) => renderBlock(block, context),
        context.comments,
      );
      if (context.comments) addComments(pages, doc.comments);
    };
    container.replaceChildren(root);
    draw();
    check();
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
        if (comments === context.comments) {
          context.changes = changes;
          root!.className = `orim-root orim-${changes}`;
          return;
        }
        context.changes = changes;
        context.comments = comments;
        draw();
      },
    };
  } catch (error) {
    const aborted = cancelled || controller.signal.aborted || renderLeases.get(container) !== lease;
    lease.cancel();
    if (aborted) throw abortError();
    throw error;
  }
}
