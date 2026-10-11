import { afterEach, describe, expect, it, vi } from "vite-plus/test";
import { StrictMode } from "react";
import { cleanup, render } from "@testing-library/react";
import * as core from "@orim/core";
import { DocxViewer } from "./DocxViewer";
import type { DocxDocument, RenderedDocument } from "@orim/core";

function createDocumentModel(text: string): DocxDocument {
  return {
    sections: [
      {
        properties: {
          width: 595,
          height: 842,
          orientation: "portrait",
          margins: { top: 48, right: 48, bottom: 48, left: 48, header: 20, footer: 20 },
        },
        blocks: [
          {
            type: "paragraph",
            properties: {},
            runs: [{ type: "text", text }],
          },
        ],
      },
    ],
    comments: [],
    imageParts: {},
  };
}

describe("DocxViewer", () => {
  afterEach(() => {
    cleanup();
    vi.restoreAllMocks();
  });

  it("renders a parsed document synchronously", () => {
    const model = createDocumentModel("React viewer");
    const renderDocument = vi.spyOn(core, "renderDocument");
    const view = render(<DocxViewer document={model} />);
    expect(view.container.textContent).toContain("React viewer");
    expect(renderDocument).toHaveBeenCalledOnce();
    expect(view.container.querySelectorAll(".orim-root")).toHaveLength(1);
  });

  it("re-renders changed options from the same model and destroys the previous render", () => {
    const model = createDocumentModel("React viewer");
    const renderDocument = vi.spyOn(core, "renderDocument");
    const view = render(<DocxViewer document={model} changes="markup" comments />);
    const previous = renderDocument.mock.results[0].value as RenderedDocument;
    const destroy = vi.spyOn(previous, "destroy");

    view.rerender(<DocxViewer document={model} changes="final" comments={false} />);

    expect(renderDocument).toHaveBeenCalledTimes(2);
    expect(renderDocument.mock.calls.map(([doc]) => doc)).toEqual([model, model]);
    expect(renderDocument.mock.calls[1][2]).toEqual({ changes: "final", comments: false });
    expect(destroy).toHaveBeenCalledOnce();
    expect(view.container.querySelector(".orim-root")?.classList.contains("orim-final")).toBe(true);
  });

  it("replaces rendered pages when the document changes", () => {
    const renderDocument = vi.spyOn(core, "renderDocument");
    const firstModel = createDocumentModel("Previous document");
    const nextModel = createDocumentModel("Latest document");
    const view = render(<DocxViewer document={firstModel} />);
    const previous = renderDocument.mock.results[0].value as RenderedDocument;
    const destroy = vi.spyOn(previous, "destroy");

    view.rerender(<DocxViewer document={nextModel} />);

    expect(destroy).toHaveBeenCalledOnce();
    expect(view.container.textContent).toContain("Latest document");
    expect(view.container.textContent).not.toContain("Previous document");
    expect(view.container.querySelectorAll(".orim-root")).toHaveLength(1);
  });

  it("empties its render container when the document becomes null", () => {
    const view = render(<DocxViewer document={createDocumentModel("Remove me")} />);
    const renderContainer = view.container.firstElementChild as HTMLElement;
    expect(renderContainer.childElementCount).toBe(1);

    view.rerender(<DocxViewer document={null} />);

    expect(renderContainer.childElementCount).toBe(0);
  });

  it("destroys the rendered document and empties its container on unmount", () => {
    const renderDocument = vi.spyOn(core, "renderDocument");
    const view = render(<DocxViewer document={createDocumentModel("Unmount me")} />);
    const rendered = renderDocument.mock.results[0].value as RenderedDocument;
    const destroy = vi.spyOn(rendered, "destroy");
    const renderContainer = view.container.firstElementChild as HTMLElement;

    view.unmount();

    expect(destroy).toHaveBeenCalledOnce();
    expect(renderContainer.childElementCount).toBe(0);
    expect(view.container.childElementCount).toBe(0);
  });

  it("leaves exactly one rendered root under StrictMode", () => {
    const view = render(
      <StrictMode>
        <DocxViewer document={createDocumentModel("Strict mode")} />
      </StrictMode>,
    );
    expect(view.container.querySelectorAll(".orim-root")).toHaveLength(1);
  });
});
