import { renderDocx, type RenderedDocument, type RenderOptions } from "../src";
import sampleUrl from "../fixtures/sample.docx?url";
import "./style.css";

const app = document.querySelector<HTMLDivElement>("#app")!;
app.innerHTML = `
  <header class="toolbar">
    <div class="brand"><span class="brand-mark">o</span><div><strong>orim</strong><small>DOCX reader</small></div></div>
    <label class="file-button">Open .docx<input id="file" type="file" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"></label>
    <div class="segmented" role="group" aria-label="Tracked changes">
      <button data-mode="markup" class="active">Markup</button><button data-mode="final">Final</button><button data-mode="original">Original</button>
    </div>
    <label class="comment-toggle"><input id="comments" type="checkbox" checked> Comments</label>
  </header>
  <main id="viewer" class="viewer"><div class="empty-state">Loading sample document…</div></main>
  <div id="drop-hint" class="drop-hint">Drop a Word document to open</div>
`;

const viewer = app.querySelector<HTMLElement>("#viewer")!;
const input = app.querySelector<HTMLInputElement>("#file")!;
const commentsToggle = app.querySelector<HTMLInputElement>("#comments")!;
const dropHint = app.querySelector<HTMLElement>("#drop-hint")!;
let activeMode: RenderOptions["changes"] = "markup";
let rendered: RenderedDocument | undefined;
let file: File | undefined;

async function showDocument(next: File | Blob): Promise<void> {
  rendered?.destroy();
  viewer.replaceChildren();
  file = next instanceof File ? next : undefined;
  try {
    rendered = await renderDocx(next, viewer, {
      changes: activeMode,
      comments: commentsToggle.checked,
    });
  } catch (error) {
    const message = document.createElement("div");
    message.className = "empty-state error";
    message.textContent = String(error instanceof Error ? error.message : error);
    viewer.replaceChildren(message);
  }
}

input.addEventListener("change", () => {
  const selected = input.files?.[0];
  if (selected) void showDocument(selected);
});

app.querySelectorAll<HTMLButtonElement>("[data-mode]").forEach((button) => {
  button.addEventListener("click", () => {
    activeMode = button.dataset.mode as RenderOptions["changes"];
    app
      .querySelectorAll("[data-mode]")
      .forEach((item) => item.classList.toggle("active", item === button));
    rendered?.setChanges(activeMode);
  });
});

commentsToggle.addEventListener("change", () => {
  if (file) void showDocument(file);
  else
    void fetch(sampleUrl)
      .then((response) => response.blob())
      .then(showDocument);
});

for (const event of ["dragenter", "dragover"] as const) {
  window.addEventListener(event, (e) => {
    e.preventDefault();
    dropHint.classList.add("visible");
  });
}
for (const event of ["dragleave", "drop"] as const) {
  window.addEventListener(event, (e) => {
    e.preventDefault();
    dropHint.classList.remove("visible");
  });
}
window.addEventListener("drop", (event) => {
  const dropped = (event as DragEvent).dataTransfer?.files[0];
  if (dropped) void showDocument(dropped);
});

void fetch(sampleUrl)
  .then((response) => response.blob())
  .then(showDocument);
