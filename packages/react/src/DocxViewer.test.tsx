import { afterEach, describe, expect, it, vi } from "vite-plus/test";
import { StrictMode } from "react";
import { cleanup, render, waitFor } from "@testing-library/react";
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
  afterEach(cleanup);

  it("updates options without reparsing the source", async () => {
    const onRender = vi.fn<(rendered: RenderedDocument) => void>();
    const model = createDocumentModel("React viewer");
    const view = render(<DocxViewer src={model} changes="markup" onRender={onRender} />);
    await view.findByText("React viewer");
    const rendered = onRender.mock.calls[0][0] as RenderedDocument;
    view.rerender(<DocxViewer src={model} changes="final" onRender={onRender} />);
    expect(onRender).toHaveBeenCalledOnce();
    expect(rendered.document).toBe(model);
    expect(view.container.querySelector(".orim-root")?.classList.contains("orim-final")).toBe(true);
    view.unmount();
  });

  it("aborts a previous source render and shows only the latest source", async () => {
    const onRender = vi.fn<(rendered: RenderedDocument) => void>();
    const onError = vi.fn();
    const view = render(<DocxViewer src={new Blob([])} onRender={onRender} onError={onError} />);
    view.rerender(
      <DocxViewer
        src={createDocumentModel("Latest source")}
        onRender={onRender}
        onError={onError}
      />,
    );
    await view.findByText("Latest source");
    expect(view.queryByText("React viewer")).toBeNull();
    expect(onRender).toHaveBeenCalledOnce();
    expect(onError).not.toHaveBeenCalled();
  });

  it("destroys the rendered document and empties its container on unmount", async () => {
    const onRender = vi.fn();
    const view = render(<DocxViewer src={createDocumentModel("Unmount me")} onRender={onRender} />);
    await view.findByText("Unmount me");
    const rendered = onRender.mock.calls[0][0] as RenderedDocument;
    const destroy = vi.spyOn(rendered, "destroy");
    view.unmount();
    expect(destroy).toHaveBeenCalledOnce();
    expect(view.container.childElementCount).toBe(0);
  });

  it("leaves exactly one rendered root under StrictMode", async () => {
    const onRender = vi.fn();
    const view = render(
      <StrictMode>
        <DocxViewer src={createDocumentModel("Strict mode")} onRender={onRender} />
      </StrictMode>,
    );
    await view.findByText("Strict mode");
    await waitFor(() => expect(onRender).toHaveBeenCalledOnce());
    expect(view.container.querySelectorAll(".orim-root")).toHaveLength(1);
    view.unmount();
  });

  it("leaves a null source empty", () => {
    const view = render(<DocxViewer src={null} />);
    expect(view.container.firstElementChild?.childElementCount ?? 0).toBe(0);
    view.unmount();
  });

  it("reports invalid DOCX bytes through onError", async () => {
    const onError = vi.fn();
    const view = render(<DocxViewer src={new Uint8Array([1, 2, 3])} onError={onError} />);
    await waitFor(() => expect(onError).toHaveBeenCalledOnce());
    view.unmount();
  });
});
