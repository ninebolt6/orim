import { afterEach, describe, expect, it, vi } from "vite-plus/test";
import { buildPptx, shape } from "../../../test/pptx";
import type { PptxPresentation, ShapeElement, Slide } from "../model";
import { parsePptx } from "../parse";
import { renderPptx, renderPresentation } from "./presentation";

const transform = { x: 10, y: 20, width: 100, height: 50, rotation: 30, flipH: true, flipV: false };
const style = { size: 20 };
const textShape = (href?: string): ShapeElement => ({
  kind: "shape",
  id: "2",
  name: "Box",
  transform,
  geometry: { preset: "ellipse", adjust: {} },
  fill: { type: "solid", color: "#ff0000" },
  line: { width: 2, fill: { type: "solid", color: "#00ff00" }, dash: "dash", tail: "triangle" },
  text: {
    anchor: "middle",
    wrap: true,
    insets: { left: 1, top: 2, right: 3, bottom: 4 },
    fontScale: 0.5,
    lineReduction: 0,
    paragraphs: [
      {
        level: 0,
        align: "center",
        marginLeft: 0,
        indent: 0,
        bullet: { text: "•" },
        endStyle: style,
        runs: [
          { text: "<b>not html</b>", style: { ...style, bold: true, color: "red;x" } },
          ...(href ? [{ text: "link", style, href }] : []),
        ],
      },
    ],
  },
});
const slide = (number: number, hidden = false, elements = [textShape()]): Slide => ({
  number,
  hidden,
  elements,
});
const presentation = (slides: readonly Slide[]): PptxPresentation => ({
  width: 400,
  height: 300,
  slides,
  media: {},
});

afterEach(() => document.body.replaceChildren());

describe("PPTX rendering", () => {
  it("positions shapes and renders sanitized text and SVG geometry", () => {
    const container = document.createElement("div");
    document.body.append(container);
    renderPresentation(
      presentation([slide(1, false, [textShape("javascript:alert(1)")]), slide(2, true)]),
      container,
    );
    const slides = container.querySelectorAll<HTMLElement>(".orim-pptx-slide");
    expect(slides).toHaveLength(1);
    expect(slides[0]!.getAttribute("aria-label")).toBe("Slide 1");
    expect(slides[0]!.style.getPropertyValue("--orim-pptx-width")).toBe("400");
    const element = container.querySelector<HTMLElement>(".orim-pptx-shape")!;
    expect([
      element.style.left,
      element.style.top,
      element.style.width,
      element.style.height,
    ]).toEqual(["2.5%", "6.6667%", "25%", "16.6667%"]);
    expect(element.style.transform).toBe("rotate(30deg)");
    const svg = element.querySelector("svg")!;
    expect(svg.style.transform).toBe("scale(-1, 1)");
    const path = svg.querySelector(":scope > path")!;
    expect(path.getAttribute("fill")).toBe("#ff0000");
    expect(path.getAttribute("stroke")).toBe("#00ff00");
    expect(path.getAttribute("stroke-dasharray")).toBe("8 6");
    expect(path.getAttribute("marker-end")).toMatch(/^url\(#orim-pptx-\d+-\d+\)$/);
    const paragraph = element.querySelector<HTMLElement>(".orim-pptx-paragraph")!;
    expect(paragraph.textContent).toBe("•<b>not html</b>link");
    expect(paragraph.querySelector("b")).toBeNull();
    expect(paragraph.querySelector("a")).toBeNull();
    const run = paragraph.querySelectorAll("span")[1]!;
    expect(run.style.fontSize).toBe("calc(var(--orim-pt) * 10)");
    expect(run.style.fontWeight).toBe("bold");
    expect(run.style.color).toBe("");
    expect(element.querySelector(".orim-pptx-text")!.className).toContain("orim-pptx-text--middle");
  });

  it("renders hidden slides on request and scrolls internal slide links", () => {
    const container = document.createElement("div");
    document.body.append(container);
    renderPresentation(
      presentation([slide(1, false, [textShape("#slide-2")]), slide(2, true)]),
      container,
      {
        hiddenSlides: true,
      },
    );
    const target = container.querySelector<HTMLElement>('[data-slide="2"]')!;
    const scroll = vi.fn();
    target.scrollIntoView = scroll;
    const link = container.querySelector("a")!;
    const event = new MouseEvent("click", { bubbles: true, cancelable: true });
    link.dispatchEvent(event);
    expect(event.defaultPrevented).toBe(true);
    expect(scroll).toHaveBeenCalled();
  });

  it("revokes media URLs on destroy and when the container is reused", async () => {
    const create = vi.spyOn(URL, "createObjectURL").mockReturnValue("blob:image");
    const revoke = vi.spyOn(URL, "revokeObjectURL").mockImplementation(() => {});
    const container = document.createElement("div");
    const deck = presentation([
      {
        number: 1,
        hidden: false,
        background: { type: "image", part: "ppt/media/a.png" },
        elements: [
          {
            kind: "picture",
            id: "3",
            name: "Picture",
            transform,
            geometry: { preset: "rect", adjust: {} },
            part: "ppt/media/a.png",
            crop: { left: 0.25, top: 0, right: 0.25, bottom: 0 },
          },
        ],
      },
    ]);
    const withMedia = {
      ...deck,
      media: { "ppt/media/a.png": { bytes: new Uint8Array([1]), mime: "image/png" } },
    };
    const first = renderPresentation(withMedia, container);
    const image = container.querySelector("img")!;
    expect(image.getAttribute("src")).toBe("blob:image");
    expect([image.style.width, image.style.left]).toEqual(["200%", "-50%"]);
    expect(create).toHaveBeenCalledTimes(1);
    const second = renderPresentation(withMedia, container);
    expect(revoke).toHaveBeenCalledTimes(1);
    second.destroy();
    expect(revoke).toHaveBeenCalledTimes(2);
    expect(container.childElementCount).toBe(0);
    first.destroy();
    create.mockRestore();
    revoke.mockRestore();
  });

  it("parses and renders a package, honouring abort signals", async () => {
    const bytes = buildPptx({
      slides: [
        {
          shapes: shape({
            id: 2,
            xfrm: [0, 0, 12700, 12700],
            body: "<a:p><a:r><a:t>Hi</a:t></a:r></a:p>",
          }),
        },
      ],
    });
    const container = document.createElement("div");
    const rendered = await renderPptx(bytes, container);
    expect(container.textContent).toBe("Hi");
    expect(rendered.presentation).toEqual(await parsePptx(bytes));
    const controller = new AbortController();
    controller.abort();
    await expect(renderPptx(bytes, container, { signal: controller.signal })).rejects.toMatchObject(
      {
        name: "AbortError",
      },
    );
    await expect(
      renderPptx(new Uint8Array([1, 2, 3]), document.createElement("div")),
    ).rejects.toThrow();
  });
});
