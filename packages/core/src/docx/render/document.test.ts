import { afterEach, describe, expect, it, vi } from "vite-plus/test";
import { renderDocx, renderDocument } from "./document";
import type { DocxDocument, Inline } from "../model";
import { splitAtPageBreaks } from "../parse/blocks";

const model = (text: string): DocxDocument => {
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
};

describe("DOCX renderer", () => {
  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("renders a parsed model synchronously", () => {
    const container = document.createElement("div");
    const rendered = renderDocument(model("synchronous"), container);
    expect(container.textContent).toContain("synchronous");
    expect(container.querySelector(".orim-root")).not.toBeNull();
    rendered.destroy();
  });

  it("destroys a synchronous render idempotently", () => {
    const container = document.createElement("div");
    const rendered = renderDocument(model("destroy"), container);
    rendered.destroy();
    rendered.destroy();
    expect(container.childElementCount).toBe(0);
  });

  it("toggles tracked-change modes and only creates safe links", async () => {
    const container = document.createElement("div");
    const rendered = renderDocument(model("deleted"), container);
    expect(container.querySelector(".orim-change--del")).not.toBeNull();
    expect(container.querySelector(".orim-comment")).not.toBeNull();
    rendered.update({ changes: "final" });
    expect(container.querySelector(".orim-root")?.classList.contains("orim-final")).toBe(true);
    expect(container.querySelector(".orim-change--del")).toBeNull();
    expect(container.textContent).toContain("inserted");
    expect(container.textContent).not.toContain("deleted");
    const links = Array.from(container.querySelectorAll("a"));
    expect(links).toHaveLength(1);
    expect(links[0].href).toBe("https://example.com/");
    expect(links[0].rel).toBe("noopener noreferrer");
    rendered.update({ changes: "original" });
    expect(container.querySelector(".orim-root")?.classList.contains("orim-original")).toBe(true);
    expect(container.textContent).toContain("deleted");
    expect(container.querySelector(".orim-change--ins")).toBeNull();
    expect(container.textContent).not.toContain("inserted");
    rendered.destroy();
    expect(container.childElementCount).toBe(0);
  });

  it("namespaces bookmark ids and rewrites internal links", () => {
    const source = model("bookmark");
    const doc: DocxDocument = {
      ...source,
      sections: source.sections.map((section) => ({
        ...section,
        blocks: [
          {
            type: "paragraph",
            properties: {},
            runs: [
              { type: "bookmark", name: "constructor" },
              { type: "text", text: "jump", href: "#constructor" },
            ],
          },
        ],
      })),
    };
    const container = document.createElement("div");
    renderDocument(doc, container);

    expect(container.querySelector("#constructor")).toBeNull();
    expect(container.querySelector("#orim-bookmark-constructor")).not.toBeNull();
    expect(container.querySelector("a")?.getAttribute("href")).toBe("#orim-bookmark-constructor");
  });

  it("keeps revision decoration visible when a run removes underline", () => {
    const source = model("text");
    const doc: DocxDocument = {
      ...source,
      sections: source.sections.map((section) => ({
        ...section,
        blocks: [
          {
            type: "paragraph",
            properties: {},
            runs: [
              {
                type: "text",
                text: "inserted without underline",
                style: { underline: "none" },
                change: { kind: "ins", author: "Aiko", date: "" },
              },
            ],
          },
        ],
      })),
    };
    const container = document.createElement("div");
    const rendered = renderDocument(doc, container);

    expect(
      container.querySelector<HTMLElement>(".orim-change--ins")?.style.textDecorationLine,
    ).toBe("");
    rendered.update({ changes: "final" });
    const plain = Array.from(container.querySelectorAll<HTMLElement>(".orim-text")).find(
      (element) => element.textContent === "inserted without underline",
    );
    expect(plain?.style.textDecorationLine).toBe("none");
    rendered.destroy();
  });

  it("renders document text as text rather than HTML", async () => {
    const container = document.createElement("div");
    renderDocument(model("<img src=x onerror=alert(1)>"), container);
    expect(container.querySelector("img")).toBeNull();
    expect(container.textContent).toContain("<img src=x onerror=alert(1)>");
  });

  it("reuses image object URLs across redraws and revokes them on destroy", () => {
    vi.spyOn(URL, "createObjectURL").mockReturnValue("blob:orim-test");
    const revoke = vi.spyOn(URL, "revokeObjectURL").mockImplementation(() => {});
    const source = model("image");
    const imageDocument: DocxDocument = {
      ...source,
      sections: source.sections.map((section) => ({
        ...section,
        blocks: [
          {
            type: "paragraph",
            properties: {},
            runs: [{ type: "image", path: "word/media/image.png", width: 1, height: 1 }],
          },
        ],
      })),
      imageParts: {
        "word/media/image.png": { bytes: new Uint8Array([1, 2, 3]), mime: "image/png" },
      },
    };
    const container = document.createElement("div");
    const rendered = renderDocument(imageDocument, container);

    expect(container.querySelector("img")?.getAttribute("src")).toBe("blob:orim-test");
    rendered.update({ changes: "final" });
    rendered.update({ comments: false });
    expect(URL.createObjectURL).toHaveBeenCalledOnce();
    rendered.destroy();
    expect(revoke).toHaveBeenCalledOnce();
    expect(revoke).toHaveBeenCalledWith("blob:orim-test");
  });

  it("filters tracked-change images by rendering mode", () => {
    vi.spyOn(URL, "createObjectURL").mockReturnValue("blob:revision-image");
    const source = model("images");
    const imageDocument: DocxDocument = {
      ...source,
      sections: source.sections.map((section) => ({
        ...section,
        blocks: [
          {
            type: "paragraph",
            properties: {},
            runs: [
              {
                type: "image",
                path: "deleted.png",
                change: { kind: "del", author: "A", date: "" },
              },
              {
                type: "image",
                path: "inserted.png",
                change: { kind: "ins", author: "B", date: "" },
              },
            ],
          },
        ],
      })),
      imageParts: {
        "deleted.png": { bytes: new Uint8Array([1]), mime: "image/png" },
        "inserted.png": { bytes: new Uint8Array([2]), mime: "image/png" },
      },
    };
    const container = document.createElement("div");
    const rendered = renderDocument(imageDocument, container);

    expect(container.querySelectorAll("img")).toHaveLength(2);
    rendered.update({ changes: "final" });
    expect(container.querySelectorAll("img")).toHaveLength(1);
    rendered.update({ changes: "original" });
    expect(container.querySelectorAll("img")).toHaveLength(1);
    rendered.destroy();
  });

  it("toggles comment balloons and anchors from a renderer model", () => {
    const container = document.createElement("div");
    const rendered = renderDocument(model("anchored text"), container);
    const anchor = container.querySelector<HTMLElement>(".orim-comment-anchor")!;
    const balloon = container.querySelector<HTMLElement>(".orim-comment")!;

    anchor.dispatchEvent(new MouseEvent("mouseover", { bubbles: true }));
    expect(anchor.classList.contains("orim-comment-active")).toBe(true);
    expect(balloon.classList.contains("orim-comment-active")).toBe(true);
    anchor.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    expect(anchor.classList.contains("orim-comment-pinned")).toBe(true);
    rendered.update({ comments: false });
    expect(container.querySelector(".orim-comment")).toBeNull();
    expect(container.querySelector(".orim-comment-anchor")).toBeNull();
    rendered.update({ comments: true });
    expect(container.querySelector(".orim-comment")).not.toBeNull();
    expect(container.querySelector(".orim-comment-anchor")).not.toBeNull();
    rendered.destroy();
  });

  it("numbers comment references by order of first appearance", () => {
    const source = model("text");
    const doc: DocxDocument = {
      ...source,
      sections: source.sections.map((section) => ({
        ...section,
        blocks: [
          {
            type: "paragraph",
            properties: {},
            runs: [
              { type: "comment-reference", id: "7" },
              { type: "comment-reference", id: "3" },
              { type: "comment-reference", id: "7" },
            ],
          },
        ],
      })),
    };
    const container = document.createElement("div");
    const rendered = renderDocument(doc, container);
    expect(
      Array.from(container.querySelectorAll("sup.orim-comment-reference")).map(
        (element) => element.textContent,
      ),
    ).toEqual(["1", "2", "1"]);
    rendered.destroy();
  });

  it("sets decoration PAGE fields at render time and resolves body fields after pagination", () => {
    const source = model("unused");
    const field = (name: string) => ({ type: "field" as const, name, cached: "cached" });
    const paragraph = (runs: Inline[]) => ({
      type: "paragraph" as const,
      properties: {},
      runs,
    });
    const doc: DocxDocument = {
      ...source,
      sections: [
        {
          properties: {
            ...source.sections[0].properties,
            header: [paragraph([field("PAGE"), { type: "text", text: "/" }, field("NUMPAGES")])],
          },
          blocks: [
            paragraph([{ type: "text", text: "first page" }]),
            {
              ...paragraph([field("PAGE")]),
              properties: { pageBreakBefore: true },
            },
          ],
        },
      ],
    };
    const container = document.createElement("div");
    const rendered = renderDocument(doc, container);
    expect(
      Array.from(container.querySelectorAll(".orim-header .orim-field")).map(
        (element) => element.textContent,
      ),
    ).toEqual(["1", "2", "2", "2"]);
    expect(container.querySelector(".orim-content .orim-field")?.textContent).toBe("2");
    rendered.destroy();
  });

  it("rejects aborted renders without leaving content behind", async () => {
    const container = document.createElement("div");
    const controller = new AbortController();
    controller.abort();
    await expect(
      renderDocx(new Blob([]), container, { signal: controller.signal }),
    ).rejects.toMatchObject({
      name: "AbortError",
    });
    expect(container.childElementCount).toBe(0);
  });

  it("prevents an older in-flight render from replacing a newer one", async () => {
    const container = document.createElement("div");
    const older = renderDocx(new Blob([]), container);
    renderDocument(model("latest"), container);
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
    renderDocument(doc, container);
    const pages = Array.from(container.querySelectorAll(".orim-page"));
    expect(pages).toHaveLength(2);
    expect(pages[0].textContent).toContain("A");
    expect(pages[1].textContent).toContain("B");
  });

  it("uses the entering section break type for continuous sections", async () => {
    const source = model("first");
    const doc: DocxDocument = {
      ...source,
      sections: [
        {
          ...source.sections[0],
          properties: { ...source.sections[0].properties, breakType: "nextPage" },
        },
        {
          ...source.sections[0],
          properties: { ...source.sections[0].properties, breakType: "continuous" },
          blocks: [
            {
              type: "paragraph",
              properties: {},
              runs: [{ type: "text", text: "second" }],
            },
          ],
        },
      ],
    };
    const container = document.createElement("div");
    renderDocument(doc, container);
    expect(container.querySelectorAll(".orim-page")).toHaveLength(1);
    expect(container.textContent).toContain("first");
    expect(container.textContent).toContain("second");
  });

  it("starts an odd-page section on the next odd page", async () => {
    const source = model("first");
    const doc: DocxDocument = {
      ...source,
      sections: [
        {
          ...source.sections[0],
          properties: { ...source.sections[0].properties, breakType: "nextPage" },
        },
        {
          ...source.sections[0],
          properties: { ...source.sections[0].properties, breakType: "oddPage" },
          blocks: [
            {
              type: "paragraph",
              properties: {},
              runs: [{ type: "text", text: "odd section" }],
            },
          ],
        },
      ],
    };
    const container = document.createElement("div");
    renderDocument(doc, container);
    expect(container.querySelectorAll(".orim-page")).toHaveLength(3);
    expect(container.querySelectorAll(".orim-page")[2].textContent).toContain("odd section");
  });

  it("sets grid pitch only when a section supplies one", async () => {
    const source = model("grid");
    const withPitch: DocxDocument = {
      ...source,
      sections: [
        {
          ...source.sections[0],
          properties: { ...source.sections[0].properties, linePitch: 18 },
        },
      ],
    };
    const pitchContainer = document.createElement("div");
    renderDocument(withPitch, pitchContainer);
    expect(
      pitchContainer
        .querySelector<HTMLElement>(".orim-content")
        ?.style.getPropertyValue("--orim-pitch"),
    ).toBe("18pt");

    const autoContainer = document.createElement("div");
    renderDocument(source, autoContainer);
    expect(
      autoContainer
        .querySelector<HTMLElement>(".orim-content")
        ?.style.getPropertyValue("--orim-pitch"),
    ).toBe("");
  });

  it("updates grid pitch when a continuous section reuses its current page", async () => {
    const source = model("grid");
    const doc: DocxDocument = {
      ...source,
      sections: [
        {
          ...source.sections[0],
          properties: { ...source.sections[0].properties, linePitch: 18 },
        },
        {
          ...source.sections[0],
          properties: {
            ...source.sections[0].properties,
            breakType: "continuous",
            linePitch: 24,
          },
          blocks: [
            {
              type: "paragraph",
              properties: {},
              runs: [{ type: "text", text: "continuous grid" }],
            },
          ],
        },
      ],
    };
    const container = document.createElement("div");
    renderDocument(doc, container);
    expect(container.querySelectorAll(".orim-page")).toHaveLength(1);
    expect(
      container.querySelector<HTMLElement>(".orim-content")?.style.getPropertyValue("--orim-pitch"),
    ).toBe("24pt");
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
    renderDocument(doc, container);
    const move = container.querySelector<HTMLElement>(".orim-change--moveTo");
    expect(move?.classList.contains("orim-change--moveTo")).toBe(true);
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
              formattingChange: { author: "Morgan", date: "" },
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
    const rendered = renderDocument(doc, container);
    rendered.update({ changes: "final" });
    expect(container.querySelectorAll(".orim-paragraph")).toHaveLength(1);
    expect(container.querySelector(".orim-paragraph-mark")).toBeNull();
    expect(container.querySelector(".orim-format-change")).toBeNull();
    expect(container.querySelector(".orim-change--moveTo")).toBeNull();

    rendered.update({ changes: "original" });
    expect(container.querySelectorAll(".orim-paragraph")).toHaveLength(2);
    expect(container.querySelector(".orim-paragraph-mark")).toBeNull();
    expect(container.querySelector(".orim-format-change")).toBeNull();
    expect(container.querySelector(".orim-change--moveFrom")).toBeNull();
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
    renderDocument(doc, container);
    expect(container.querySelectorAll(".orim-page-comments > .orim-comment")).toHaveLength(1);
    expect(container.querySelectorAll(".orim-comment-replies .orim-comment")).toHaveLength(1);
    expect(container.querySelector(".orim-comment-author")?.textContent).toContain("Aiko · ");
  });

  it("renders paragraph layout and table formatting from an immutable model", () => {
    const source = model("formatted");
    const doc: DocxDocument = {
      ...source,
      sections: source.sections.map((section) => ({
        ...section,
        blocks: [
          {
            type: "paragraph",
            properties: {
              styleId: "Heading 1",
              align: "both",
              before: 2,
              after: 3,
              line: 1.5,
              lineRule: "exact",
              left: 4,
              right: 5,
              firstLine: 6,
              hanging: 7,
              pageBreakBefore: true,
              snapToGrid: false,
              shading: "#ffeecc",
              borders: { bottom: "1px solid red" },
              borderSpacing: { top: 2 },
              change: { kind: "ins", author: "Aiko", date: "2025-01-01" },
              formattingChange: { author: "Morgan", date: "2025-01-02" },
            },
            runs: [{ type: "text", text: "formatted paragraph" }],
            markStyle: { bold: true },
            list: { numId: "1", level: 0, marker: "1.", left: 12, hanging: 3, suffix: "space" },
          },
          {
            type: "table",
            width: "100%",
            align: "center",
            widths: ["50%", "50%"],
            borders: { insideV: "1px solid blue" },
            margins: { top: 2 },
            rows: [
              {
                header: true,
                height: 18,
                change: { kind: "ins", author: "Aiko", date: "" },
                cells: [
                  {
                    blocks: [
                      { type: "paragraph", properties: {}, runs: [{ type: "text", text: "cell" }] },
                    ],
                    gridSpan: 2,
                    rowSpan: 2,
                    width: "50%",
                    shading: "#eeeeee",
                    borders: { left: "2px solid green" },
                    margins: { right: 4 },
                    verticalAlign: "middle",
                  },
                  {
                    blocks: [],
                    gridSpan: 1,
                    rowSpan: 0,
                  },
                ],
              },
              {
                change: { kind: "del", author: "Morgan", date: "" },
                cells: [
                  {
                    blocks: [],
                    gridSpan: 1,
                    rowSpan: 1,
                  },
                ],
              },
            ],
          },
        ],
      })),
    };
    const container = document.createElement("div");
    const rendered = renderDocument(doc, container);
    const paragraph = container.querySelector<HTMLElement>(".orim-paragraph")!;
    const table = container.querySelector<HTMLTableElement>(".orim-table")!;
    const cell = container.querySelector<HTMLTableCellElement>(".orim-table-cell")!;

    expect(paragraph.style.textAlign).toBe("justify");
    expect(paragraph.style.marginTop).toBe("2pt");
    expect(paragraph.style.lineHeight).toBe("1.5pt");
    expect(paragraph.style.breakBefore).toBe("page");
    expect(paragraph.classList.contains("orim-format-change")).toBe(true);
    expect(paragraph.querySelector(".orim-list-marker")?.textContent).toBe("1. ");
    expect(table.style.width).toBe("100%");
    expect(table.style.marginInline).toBe("auto");
    expect(table.querySelectorAll("col")).toHaveLength(2);
    expect(table.querySelector("tr")?.dataset.orimHeader).toBe("true");
    expect(table.querySelectorAll("tr")).toHaveLength(2);
    expect(cell.colSpan).toBe(2);
    expect(cell.rowSpan).toBe(2);
    expect(cell.style.verticalAlign).toBe("middle");
    expect(cell.style.paddingRight).toBe("4pt");

    rendered.update({ changes: "final" });
    expect(container.querySelectorAll(".orim-table tr")).toHaveLength(1);
    rendered.destroy();
  });
});
