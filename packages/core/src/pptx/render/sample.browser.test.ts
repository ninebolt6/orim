import { describe, expect, it } from "vite-plus/test";
import { parsePptx } from "../parse";
import { renderPresentation } from "./presentation";
import "./styles.css";
import sampleUrl from "../../../../../fixtures/sample.pptx?url";

const samplePresentation = async () => parsePptx(await (await fetch(sampleUrl)).blob());

describe("sample PPTX in Chromium", () => {
  it("scales slides to the container width and keeps geometry proportional", async () => {
    const presentation = await samplePresentation();
    const container = document.createElement("div");
    container.style.width = "960px";
    document.body.append(container);
    const rendered = renderPresentation(presentation, container);
    const slides = container.querySelectorAll<HTMLElement>(".orim-pptx-slide");
    expect(slides).toHaveLength(4);
    const box = slides[0]!.getBoundingClientRect();
    expect(box.width).toBeCloseTo(960, 0);
    expect(box.height).toBeCloseTo(540, 0);
    const title = [...slides[1]!.querySelectorAll<HTMLElement>(".orim-pptx-shape")].find(
      (element) => element.textContent === "Text and lists",
    )!;
    const titleBox = title.getBoundingClientRect();
    const slideBox = slides[1]!.getBoundingClientRect();
    expect(titleBox.left - slideBox.left).toBeCloseTo(48, 0);
    expect(titleBox.width).toBeCloseTo(864, 0);
    const titleRun = title.querySelector<HTMLElement>(".orim-pptx-paragraph span")!;
    expect(Number.parseFloat(getComputedStyle(titleRun).fontSize)).toBeCloseTo(36, 0);
    expect(getComputedStyle(titleRun).fontWeight).toBe("700");

    container.style.width = "480px";
    expect(Number.parseFloat(getComputedStyle(titleRun).fontSize)).toBeCloseTo(18, 0);
    expect(title.getBoundingClientRect().width).toBeCloseTo(432, 0);

    const picture = slides[2]!.querySelector<HTMLImageElement>(".orim-pptx-picture img")!;
    await picture.decode();
    expect(picture.naturalWidth).toBe(64);
    const frame = picture.parentElement!.getBoundingClientRect();
    expect(picture.getBoundingClientRect().width).toBeCloseTo(frame.width / 0.8, 0);
    const numbers = [...container.querySelectorAll<HTMLElement>('[data-name="Slide Number"]')].map(
      (element) => element.textContent,
    );
    expect(numbers).toEqual(["2", "3", "4"]);
    rendered.destroy();
    container.remove();
  });
});
