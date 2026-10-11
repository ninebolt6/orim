import { describe, expect, it } from "vite-plus/test";
import { renderDocx } from "./document";
import type { DocxDocument } from "../docx/model";
import { splitAtPageBreaks } from "../docx/parse/blocks";

function model(text: string): DocxDocument {
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
            runs: [
              {
                type: "text",
                text,
                change: { kind: "del", author: "Aiko", date: "" },
                comments: ["0"],
              },
              {
                type: "text",
                text: "inserted",
                change: { kind: "ins", author: "Morgan", date: "" },
              },
              { type: "text", text: "unsafe", href: "javascript:alert(1)" },
              { type: "text", text: "safe", href: "https://example.com" },
            ],
          },
        ],
      },
    ],
    comments: [
      {
        id: "0",
        author: "Aiko",
        paragraphs: [
          {
            type: "paragraph",
            properties: {},
            runs: [{ type: "text", text: "Review this passage." }],
          },
        ],
      },
    ],
    imageParts: {},
  };
}

describe("DOCX renderer", () => {
  it("toggles tracked-change modes and only creates safe links", async () => {
    const container = document.createElement("div");
    const rendered = await renderDocx(model("deleted"), container);
    expect(container.querySelector(".orim-change--del")).not.toBeNull();
    expect(container.querySelector(".orim-comment")).not.toBeNull();
    rendered.update({ changes: "final" });
    expect(container.querySelector(".orim-root")?.classList.contains("orim-final")).toBe(true);
    expect(container.querySelector(".orim-change--del")).not.toBeNull();
    expect(container.textContent).toContain("inserted");
    const links = Array.from(container.querySelectorAll("a"));
    expect(links).toHaveLength(1);
    expect(links[0].href).toBe("https://example.com/");
    rendered.update({ changes: "original" });
    expect(container.querySelector(".orim-root")?.classList.contains("orim-original")).toBe(true);
    expect(container.textContent).toContain("deleted");
    expect(container.querySelector(".orim-change--ins")).not.toBeNull();
    rendered.destroy();
    expect(container.childElementCount).toBe(0);
  });

  it("renders document text as text rather than HTML", async () => {
    const container = document.createElement("div");
    await renderDocx(model("<img src=x onerror=alert(1)>"), container);
    expect(container.querySelector("img")).toBeNull();
    expect(container.textContent).toContain("<img src=x onerror=alert(1)>");
  });

  it("rejects aborted renders without leaving content behind", async () => {
    const container = document.createElement("div");
    const controller = new AbortController();
    controller.abort();
    await expect(
      renderDocx(model("aborted"), container, { signal: controller.signal }),
    ).rejects.toMatchObject({
      name: "AbortError",
    });
    expect(container.childElementCount).toBe(0);
  });

  it("prevents an older in-flight render from replacing a newer one", async () => {
    const container = document.createElement("div");
    const older = renderDocx(new Blob([]), container);
    await renderDocx(model("latest"), container);
    await expect(older).rejects.toMatchObject({ name: "AbortError" });
    expect(container.textContent).toContain("latest");
    expect(container.textContent).not.toContain("old");
  });

  it("does not add a blank page before a paragraph after an explicit page break", async () => {
    const source = model("unused");
    const doc: DocxDocument = {
      ...source,
      sections: source.sections.map((section) => ({
        ...section,
        blocks: [
          ...splitAtPageBreaks({
            type: "paragraph",
            properties: {},
            runs: [
              { type: "text", text: "A" },
              { type: "break", page: true },
            ],
          }),
          {
            type: "paragraph",
            properties: { pageBreakBefore: true },
            runs: [{ type: "text", text: "B" }],
          },
        ],
      })),
    };
    const container = document.createElement("div");
    await renderDocx(doc, container);
    const pages = Array.from(container.querySelectorAll(".orim-page"));
    expect(pages).toHaveLength(2);
    expect(pages[0].textContent).toContain("A");
    expect(pages[1].textContent).toContain("B");
  });

  it("keeps moved text green regardless of its run color", async () => {
    const source = model("moved");
    const doc: DocxDocument = {
      ...source,
      sections: source.sections.map((section) => ({
        ...section,
        blocks: section.blocks.map((block) =>
          block.type === "paragraph"
            ? {
                ...block,
                runs: [
                  ...block.runs,
                  {
                    type: "text",
                    text: "move",
                    style: { color: "#ff0000" },
                    change: { kind: "moveTo", author: "Aiko", date: "" },
                  },
                ],
              }
            : block,
        ),
      })),
    };
    const container = document.createElement("div");
    await renderDocx(doc, container);
    const move = container.querySelector<HTMLElement>(".orim-change--moveTo");
    expect(move?.style.color).not.toBe("#ff0000");
    expect(move?.style.textDecoration).toBe("underline double");
  });

  it("renders accepted moves plainly and removes revision decoration outside Markup", async () => {
    const source = model("text");
    const doc: DocxDocument = {
      ...source,
      sections: source.sections.map((section) => ({
        ...section,
        blocks: [
          {
            type: "paragraph",
            properties: { change: { kind: "del", author: "Aiko", date: "" } },
            runs: [{ type: "text", text: "deleted paragraph mark" }],
          },
          {
            type: "paragraph",
            properties: {
              formattingChange: { kind: "ins", author: "Morgan", date: "" },
              change: { kind: "ins", author: "Morgan", date: "" },
            },
            runs: [
              {
                type: "text",
                text: "move to",
                change: { kind: "moveTo", author: "Morgan", date: "" },
              },
              {
                type: "text",
                text: "move from",
                change: { kind: "moveFrom", author: "Aiko", date: "" },
              },
            ],
          },
        ],
      })),
    };
    const container = document.createElement("div");
    const rendered = await renderDocx(doc, container);
    rendered.update({ changes: "final" });
    expect(container.querySelectorAll(".orim-paragraph")).toHaveLength(1);
    expect(container.querySelector(".orim-paragraph-mark")).toBeNull();
    expect(container.querySelector(".orim-format-change")).toBeNull();
    expect(container.querySelector<HTMLElement>(".orim-change--moveTo")?.style.color).toBe(
      "inherit",
    );
    expect(container.querySelector<HTMLElement>(".orim-change--moveTo")?.style.textDecoration).toBe(
      "none",
    );

    rendered.update({ changes: "original" });
    expect(container.querySelectorAll(".orim-paragraph")).toHaveLength(2);
    expect(container.querySelector(".orim-paragraph-mark")).toBeNull();
    expect(container.querySelector(".orim-format-change")).toBeNull();
    expect(container.querySelector<HTMLElement>(".orim-change--moveFrom")?.style.color).toBe(
      "inherit",
    );
    expect(
      container.querySelector<HTMLElement>(".orim-change--moveFrom")?.style.textDecoration,
    ).toBe("none");
    rendered.destroy();
  });

  it("nests comment replies without rendering a second top-level balloon", async () => {
    const source = model("anchored text");
    const doc: DocxDocument = {
      ...source,
      comments: [
        { ...source.comments[0], date: "2025-02-01T09:15:00Z" },
        {
          id: "1",
          parentId: "0",
          author: "Morgan",
          paragraphs: [],
        },
      ],
    };
    const container = document.createElement("div");
    await renderDocx(doc, container);
    expect(container.querySelectorAll(".orim-page-comments > .orim-comment")).toHaveLength(1);
    expect(container.querySelectorAll(".orim-comment-replies .orim-comment")).toHaveLength(1);
    expect(container.querySelector(".orim-comment-author")?.textContent).toContain("Aiko · ");
  });
});
