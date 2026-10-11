import { afterEach, describe, expect, it, vi } from "vite-plus/test";
import { cleanup, render, waitFor } from "@testing-library/react";
import { DocxViewer } from "./DocxViewer";
import type { DocxDocument } from "@orim/core";

const documentModel: DocxDocument = {
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
          runs: [{ type: "text", text: "React viewer" }],
        },
      ],
    },
  ],
  comments: [],
  imageParts: {},
};

describe("DocxViewer", () => {
  afterEach(cleanup);

  it("updates options without replacing the rendered page nodes", async () => {
    const view = render(<DocxViewer src={documentModel} changes="markup" />);
    await view.findByText("React viewer");
    const page = view.container.querySelector(".orim-page");
    view.rerender(<DocxViewer src={documentModel} changes="final" />);
    expect(view.container.querySelector(".orim-page")).toBe(page);
    expect(view.container.querySelector(".orim-root")?.classList.contains("orim-final")).toBe(true);
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
