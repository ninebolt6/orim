import type { DocxComment } from "../docx/model";
import type { Page } from "./paginate";

function renderCommentText(comment: DocxComment): string {
  return comment.paragraphs
    .map((block) =>
      block.type === "paragraph"
        ? block.runs.map((run) => (run.type === "text" ? run.text : "")).join("")
        : "",
    )
    .join("\n");
}

function renderCommentAuthor(comment: DocxComment): string {
  if (!comment.date || !Number.isFinite(new Date(comment.date).getTime())) return comment.author;
  return `${comment.author} · ${new Date(comment.date).toLocaleString(undefined, {
    dateStyle: "medium",
    timeStyle: "short",
  })}`;
}

export function addComments(
  pages: readonly Page[],
  comments: readonly DocxComment[],
  anchors: ReadonlyMap<string, readonly HTMLElement[]>,
  balloons: Map<string, HTMLElement>,
  scope: HTMLElement,
): () => void {
  const byId = new Map(comments.map((comment) => [comment.id, comment]));
  for (const page of pages) {
    const lane = page.wrapper.querySelector<HTMLElement>(".orim-page-comments");
    if (!lane) continue;
    const ordered = new Map<string, number>();
    for (const [id, elements] of anchors) {
      for (const anchor of elements) {
        if (!page.content.contains(anchor)) continue;
        const top = anchor.getBoundingClientRect().top - page.page.getBoundingClientRect().top;
        if (!ordered.has(id)) ordered.set(id, top);
      }
    }
    let nextTop = 0;
    for (const [id, top] of [...ordered].sort((a, b) => a[1] - b[1])) {
      const comment = byId.get(id);
      if (!comment || comment.parentId) continue;
      const balloon = document.createElement("div");
      balloon.className = "orim-comment";
      balloon.dataset.comment = id;
      balloons.set(id, balloon);
      balloon.style.top = `${Math.max(top, nextTop)}px`;
      if (comment.done) balloon.classList.add("orim-comment--done");
      const author = document.createElement("div");
      author.className = "orim-comment-author";
      author.textContent = renderCommentAuthor(comment);
      balloon.append(author);
      const text = document.createElement("div");
      text.className = "orim-comment-text";
      text.textContent = renderCommentText(comment);
      balloon.append(text);
      const replies = comments.filter((reply) => reply.parentId === comment.id);
      if (replies.length) {
        const replyBox = document.createElement("div");
        replyBox.className = "orim-comment-replies";
        for (const reply of replies) {
          const replyEl = document.createElement("div");
          replyEl.className = `orim-comment${reply.done ? " orim-comment--done" : ""}`;
          const replyAuthor = document.createElement("div");
          replyAuthor.className = "orim-comment-author";
          replyAuthor.textContent = renderCommentAuthor(reply);
          const replyText = document.createElement("div");
          replyText.textContent = renderCommentText(reply);
          replyEl.append(replyAuthor, replyText);
          replyBox.append(replyEl);
        }
        balloon.append(replyBox);
      }
      lane.append(balloon);
      nextTop = Number.parseFloat(balloon.style.top) + balloon.offsetHeight + 8;
    }
  }

  const highlight = (id: string | undefined, active: boolean) => {
    if (!id) return;
    const balloon = balloons.get(id);
    const pinned =
      balloon?.classList.contains("orim-comment-pinned") ||
      anchors.get(id)?.some((anchor) => anchor.classList.contains("orim-comment-pinned"));
    for (const anchor of anchors.get(id) ?? []) {
      anchor.classList.toggle("orim-comment-active", active || Boolean(pinned));
    }
    balloon?.classList.toggle("orim-comment-active", active || Boolean(pinned));
  };
  const togglePinned = (id: string | undefined) => {
    if (!id) return;
    const balloon = balloons.get(id);
    const active = !balloon?.classList.contains("orim-comment-pinned");
    for (const anchor of anchors.get(id) ?? [])
      anchor.classList.toggle("orim-comment-pinned", active);
    balloon?.classList.toggle("orim-comment-pinned", active);
    highlight(id, active);
  };
  const targetFor = (event: Event): HTMLElement | undefined => {
    if (!(event.target instanceof Element)) return undefined;
    return (
      event.target.closest<HTMLElement>(
        ".orim-comment[data-comment],.orim-comment-anchor[data-c]",
      ) ?? undefined
    );
  };
  const commentId = (element: HTMLElement): string | undefined =>
    element.dataset.comment ?? element.dataset.c?.split(" ")[0];
  const onMouseOver = (event: MouseEvent) => {
    const target = targetFor(event);
    if (target && !(event.relatedTarget instanceof Node && target.contains(event.relatedTarget)))
      highlight(commentId(target), true);
  };
  const onMouseOut = (event: MouseEvent) => {
    const target = targetFor(event);
    if (target && !(event.relatedTarget instanceof Node && target.contains(event.relatedTarget)))
      highlight(commentId(target), false);
  };
  const onClick = (event: MouseEvent) => {
    const target = targetFor(event);
    if (target) togglePinned(commentId(target));
  };
  scope.addEventListener("mouseover", onMouseOver);
  scope.addEventListener("mouseout", onMouseOut);
  scope.addEventListener("click", onClick);
  return () => {
    scope.removeEventListener("mouseover", onMouseOver);
    scope.removeEventListener("mouseout", onMouseOut);
    scope.removeEventListener("click", onClick);
  };
}
