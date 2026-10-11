import { describe, expect, it } from "vite-plus/test";
import { renderDocx } from "./render";
import type { DocxDocument } from "../docx/model";

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
    rendered.setChanges("final");
    expect(container.firstElementChild?.classList.contains("orim-final")).toBe(true);
    expect(container.textContent).not.toContain("deleted");
    expect(container.textContent).toContain("inserted");
    const links = Array.from(container.querySelectorAll("a"));
    expect(links).toHaveLength(1);
    expect(links[0].href).toBe("https://example.com/");
    rendered.setChanges("original");
    expect(container.firstElementChild?.classList.contains("orim-original")).toBe(true);
    expect(container.textContent).toContain("deleted");
    expect(container.textContent).not.toContain("inserted");
    rendered.destroy();
    expect(container.childElementCount).toBe(0);
  });

  it("renders document text as text rather than HTML", async () => {
    const container = document.createElement("div");
    await renderDocx(model("<img src=x onerror=alert(1)>"), container);
    expect(container.querySelector("img")).toBeNull();
    expect(container.textContent).toContain("<img src=x onerror=alert(1)>");
  });

  it("keeps moved text green regardless of its run color", async () => {
    const doc = model("moved");
    const paragraph = doc.sections[0].blocks[0];
    if (paragraph.type === "paragraph")
      paragraph.runs.push({
        type: "text",
        text: "move",
        style: { color: "#ff0000" },
        change: { kind: "moveTo", author: "Aiko", date: "" },
      });
    const container = document.createElement("div");
    await renderDocx(doc, container);
    const move = container.querySelector<HTMLElement>(".orim-change--moveTo");
    expect(move?.style.color).not.toBe("#ff0000");
    expect(move?.style.textDecoration).toBe("underline double");
  });

  it("nests comment replies without rendering a second top-level balloon", async () => {
    const doc = model("anchored text");
    doc.comments[0].date = "2025-02-01T09:15:00Z";
    doc.comments.push({
      id: "1",
      parentId: "0",
      author: "Morgan",
      paragraphs: [],
    });
    const container = document.createElement("div");
    await renderDocx(doc, container);
    expect(container.querySelectorAll(".orim-page-comments > .orim-comment")).toHaveLength(1);
    expect(container.querySelectorAll(".orim-comment-replies .orim-comment")).toHaveLength(1);
    expect(container.querySelector(".orim-comment-author")?.textContent).toContain("Aiko · ");
  });
});
