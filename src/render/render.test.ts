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
});
