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

export function addComments(pages: Page[], comments: DocxComment[]): void {
  const scope = pages[0]?.wrapper.parentElement;
  if (!scope) return;
  const byId = new Map(comments.map((comment) => [comment.id, comment]));
  for (const page of pages) {
    const lane = page.wrapper.querySelector<HTMLElement>(".orim-page-comments");
    if (!lane) continue;
    const candidates = Array.from(
      page.content.querySelectorAll<HTMLElement>(".orim-comment-anchor[data-c]"),
    );
    const ordered = new Map<string, number>();
    for (const anchor of candidates) {
      const top = anchor.getBoundingClientRect().top - page.page.getBoundingClientRect().top;
      for (const id of anchor.dataset.c?.split(" ") ?? []) {
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
    for (const anchor of scope.querySelectorAll<HTMLElement>(".orim-comment-anchor[data-c]")) {
      if (anchor.dataset.c?.split(" ").includes(id))
        anchor.classList.toggle(
          "orim-comment-active",
          active || anchor.classList.contains("orim-comment-pinned"),
        );
    }
    for (const balloon of scope.querySelectorAll<HTMLElement>(".orim-comment[data-comment]")) {
      if (balloon.dataset.comment === id)
        balloon.classList.toggle(
          "orim-comment-active",
          active || balloon.classList.contains("orim-comment-pinned"),
        );
    }
  };
  const togglePinned = (id: string | undefined) => {
    if (!id) return;
    const balloon = scope.querySelector<HTMLElement>(
      `.orim-comment[data-comment="${CSS.escape(id)}"]`,
    );
    const active = !balloon?.classList.contains("orim-comment-pinned");
    for (const element of scope.querySelectorAll<HTMLElement>(
      `.orim-comment[data-comment="${CSS.escape(id)}"],.orim-comment-anchor[data-c]`,
    )) {
      if (
        element.classList.contains("orim-comment-anchor") &&
        !element.dataset.c?.split(" ").includes(id)
      )
        continue;
      element.classList.toggle("orim-comment-pinned", active);
    }
    highlight(id, active);
  };
  for (const balloon of scope.querySelectorAll<HTMLElement>(".orim-comment[data-comment]")) {
    balloon.addEventListener("mouseenter", () => highlight(balloon.dataset.comment, true));
    balloon.addEventListener("mouseleave", () => highlight(balloon.dataset.comment, false));
    balloon.addEventListener("click", () => togglePinned(balloon.dataset.comment));
  }
  for (const anchor of scope.querySelectorAll<HTMLElement>(".orim-comment-anchor")) {
    const id = anchor.dataset.c?.split(" ")[0];
    anchor.addEventListener("mouseenter", () => highlight(id, true));
    anchor.addEventListener("mouseleave", () => highlight(id, false));
    anchor.addEventListener("click", () => togglePinned(id));
  }
}
