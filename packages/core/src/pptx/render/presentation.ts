import { acquireContainerLease } from "../../common/lease";
import { createObjectUrls } from "../../common/media";
import { renderParsed } from "../../common/render-file";
import { safeHref, validColor } from "../../common/sanitize";
import type { PackageInput } from "../../opc/package";
import type {
  Bounds,
  Fill,
  LineStyle,
  PictureElement,
  PptxPresentation,
  ShapeElement,
  SlideElement,
  TableElement,
} from "../model";
import { parsePptx } from "../parse";
import { geometryPaths } from "./geometry";
import { border, fillBackground, lineColor, pt, type RenderContext } from "./style";
import { renderParagraphs, renderTextBody } from "./text";

export interface RenderPresentationOptions {
  /** Render slides marked hidden in the deck. Defaults to `false`. */
  readonly hiddenSlides?: boolean;
}

export interface RenderPptxOptions extends RenderPresentationOptions {
  readonly signal?: AbortSignal;
}

export interface RenderedPresentation {
  readonly presentation: PptxPresentation;
  destroy(): void;
}

const SVG = "http://www.w3.org/2000/svg";
let renderSequence = 0;

const percent = (value: number): string => `${Number((value * 100).toFixed(4))}%`;

const svg = <K extends keyof SVGElementTagNameMap>(
  name: K,
  attributes: Readonly<Record<string, string | number>>,
): SVGElementTagNameMap[K] => {
  const element = document.createElementNS(SVG, name);
  for (const [key, value] of Object.entries(attributes)) element.setAttribute(key, String(value));
  return element;
};

const dashes: Readonly<Record<string, readonly number[]>> = {
  dash: [4, 3],
  sysDash: [3, 1],
  dot: [1, 3],
  sysDot: [1, 1],
  dashDot: [4, 3, 1, 3],
  sysDashDot: [3, 1, 1, 1],
  lgDash: [8, 3],
  lgDashDot: [8, 3, 1, 3],
};

const svgPaint = (
  fill: Fill | undefined,
  defs: SVGDefsElement,
  bounds: Bounds,
  context: RenderContext,
): string => {
  switch (fill?.type) {
    case "solid":
      return validColor(fill.color) ?? "none";
    case "gradient": {
      const id = context.id();
      const radians = (fill.angle * Math.PI) / 180;
      const gradient = fill.radial
        ? svg("radialGradient", { id })
        : svg("linearGradient", {
            id,
            x1: 0.5 - Math.cos(radians) / 2,
            y1: 0.5 - Math.sin(radians) / 2,
            x2: 0.5 + Math.cos(radians) / 2,
            y2: 0.5 + Math.sin(radians) / 2,
          });
      for (const stop of fill.stops)
        gradient.append(
          svg("stop", { offset: stop.position, "stop-color": validColor(stop.color) ?? "#000000" }),
        );
      defs.append(gradient);
      return `url(#${id})`;
    }
    case "image": {
      const url = context.urls.url(fill.part, context.presentation.media[fill.part]);
      if (!url) return "none";
      const id = context.id();
      const pattern = svg("pattern", {
        id,
        patternUnits: "userSpaceOnUse",
        width: bounds.width,
        height: bounds.height,
      });
      pattern.append(
        svg("image", {
          href: url,
          width: bounds.width,
          height: bounds.height,
          preserveAspectRatio: "none",
        }),
      );
      defs.append(pattern);
      return `url(#${id})`;
    }
  }
  return "none";
};

const marker = (
  type: string,
  color: string,
  defs: SVGDefsElement,
  context: RenderContext,
): string => {
  const id = context.id();
  const element = svg("marker", {
    id,
    viewBox: "0 0 10 10",
    refX: type === "oval" ? 5 : 8,
    refY: 5,
    markerWidth: 3,
    markerHeight: 3,
    orient: "auto-start-reverse",
  });
  element.append(
    type === "oval"
      ? svg("circle", { cx: 5, cy: 5, r: 5, fill: color })
      : svg("path", {
          d: type === "diamond" ? "M0 5L5 0L10 5L5 10Z" : "M0 0L10 5L0 10Z",
          fill: color,
        }),
  );
  defs.append(element);
  return `url(#${id})`;
};

const renderGeometry = (
  shape: ShapeElement | PictureElement,
  fill: Fill | undefined,
  line: LineStyle | undefined,
  context: RenderContext,
): SVGSVGElement => {
  const { width, height } = shape.transform;
  const root = svg("svg", {
    class: "orim-pptx-geometry",
    viewBox: `0 0 ${Math.max(width, 0.01)} ${Math.max(height, 0.01)}`,
    preserveAspectRatio: "none",
    "aria-hidden": "true",
  });
  const defs = svg("defs", {});
  root.append(defs);
  const paint = svgPaint(fill, defs, shape.transform, context);
  const stroke = lineColor(line);
  const strokeWidth = line?.width ?? 0.75;
  for (const path of geometryPaths(shape.geometry, width, height)) {
    const element = svg("path", {
      d: path.data,
      fill: path.fill ? paint : "none",
      stroke: path.stroke && stroke ? stroke : "none",
      "stroke-width": strokeWidth,
      "stroke-linejoin": "round",
    });
    const dash = line?.dash && dashes[line.dash];
    if (dash)
      element.setAttribute("stroke-dasharray", dash.map((item) => item * strokeWidth).join(" "));
    if (stroke && line?.head)
      element.setAttribute("marker-start", marker(line.head, stroke, defs, context));
    if (stroke && line?.tail)
      element.setAttribute("marker-end", marker(line.tail, stroke, defs, context));
    root.append(element);
  }
  return root;
};

const flip = (element: Element & ElementCSSInlineStyle, x: boolean, y: boolean): void => {
  if (x || y) element.style.transform = `scale(${x ? -1 : 1}, ${y ? -1 : 1})`;
};

const renderPicture = (
  picture: PictureElement,
  element: HTMLElement,
  context: RenderContext,
): void => {
  const frame = document.createElement("div");
  frame.className = "orim-pptx-picture";
  if (picture.geometry.preset === "ellipse") frame.style.clipPath = "ellipse(50% 50%)";
  flip(frame, picture.transform.flipH, picture.transform.flipV);
  const image = document.createElement("img");
  image.alt = picture.description ?? "";
  const url = context.urls.url(picture.part, context.presentation.media[picture.part]);
  if (url) image.src = url;
  const crop = picture.crop ?? { left: 0, top: 0, right: 0, bottom: 0 };
  const width = 1 - crop.left - crop.right || 1;
  const height = 1 - crop.top - crop.bottom || 1;
  image.style.width = percent(1 / width);
  image.style.height = percent(1 / height);
  image.style.left = percent(-crop.left / width);
  image.style.top = percent(-crop.top / height);
  frame.append(image);
  element.append(frame);
  if (lineColor(picture.line))
    element.append(renderGeometry(picture, undefined, picture.line, context));
};

const renderTable = (table: TableElement, element: HTMLElement, context: RenderContext): void => {
  const total = table.columns.reduce((sum, width) => sum + width, 0) || 1;
  const grid = document.createElement("table");
  grid.className = "orim-pptx-table";
  const columns = document.createElement("colgroup");
  for (const width of table.columns) {
    const column = document.createElement("col");
    column.style.width = percent(width / total);
    columns.append(column);
  }
  const body = document.createElement("tbody");
  for (const row of table.rows) {
    const tableRow = document.createElement("tr");
    tableRow.style.height = pt(row.height);
    for (const cell of row.cells) {
      if (cell.merged) continue;
      const tableCell = document.createElement("td");
      if (cell.columnSpan > 1) tableCell.colSpan = cell.columnSpan;
      if (cell.rowSpan > 1) tableCell.rowSpan = cell.rowSpan;
      const background = fillBackground(cell.fill, context);
      if (background) tableCell.style.background = background;
      tableCell.style.borderLeft = border(cell.borders.left);
      tableCell.style.borderRight = border(cell.borders.right);
      tableCell.style.borderTop = border(cell.borders.top);
      tableCell.style.borderBottom = border(cell.borders.bottom);
      if (cell.text) {
        const { insets, anchor } = cell.text;
        tableCell.style.padding = [insets.top, insets.right, insets.bottom, insets.left]
          .map(pt)
          .join(" ");
        tableCell.style.verticalAlign = anchor;
        renderParagraphs(cell.text, tableCell);
      }
      tableRow.append(tableCell);
    }
    body.append(tableRow);
  }
  grid.append(columns, body);
  element.append(grid);
};

const renderElement = (item: SlideElement, frame: Bounds, context: RenderContext): HTMLElement => {
  const href = safeHref(item.href);
  const element = document.createElement(href ? "a" : "div");
  if (href && element instanceof HTMLAnchorElement) {
    element.href = href;
    if (!href.startsWith("#")) element.rel = "noopener noreferrer";
  }
  element.className = `orim-pptx-element orim-pptx-${item.kind}`;
  if (item.name) element.dataset.name = item.name;
  const { x, y, width, height, rotation } = item.transform;
  element.style.left = percent(frame.width ? (x - frame.x) / frame.width : 0);
  element.style.top = percent(frame.height ? (y - frame.y) / frame.height : 0);
  element.style.width = percent(frame.width ? width / frame.width : 0);
  element.style.height = percent(frame.height ? height / frame.height : 0);
  if (rotation) element.style.transform = `rotate(${Number(rotation.toFixed(3))}deg)`;
  switch (item.kind) {
    case "shape": {
      const geometry = renderGeometry(item, item.fill, item.line, context);
      flip(geometry, item.transform.flipH, item.transform.flipV);
      element.append(geometry);
      if (item.text?.paragraphs.some((paragraph) => paragraph.runs.length))
        element.append(renderTextBody(item.text));
      break;
    }
    case "picture":
      renderPicture(item, element, context);
      break;
    case "group": {
      const scale = item.transform.flipH || item.transform.flipV;
      if (scale)
        element.style.transform = `${element.style.transform} scale(${item.transform.flipH ? -1 : 1}, ${item.transform.flipV ? -1 : 1})`;
      for (const child of item.children)
        element.append(renderElement(child, item.childBounds, context));
      break;
    }
    case "table":
      renderTable(item, element, context);
      break;
  }
  return element;
};

export const renderPresentation = (
  presentation: PptxPresentation,
  container: HTMLElement,
  options: RenderPresentationOptions = {},
): RenderedPresentation => {
  const urls = createObjectUrls();
  let root: HTMLDivElement | undefined;
  const lease = acquireContainerLease(container, () => {
    if (root?.parentElement === container) container.replaceChildren();
    urls.revoke();
  });
  try {
    const prefix = `orim-pptx-${++renderSequence}-`;
    let sequence = 0;
    const context: RenderContext = { presentation, urls, id: () => `${prefix}${++sequence}` };
    root = document.createElement("div");
    root.className = "orim-pptx";
    const frame: Bounds = { x: 0, y: 0, width: presentation.width, height: presentation.height };
    const sections = new Map<number, HTMLElement>();
    for (const slide of presentation.slides) {
      if (slide.hidden && !options.hiddenSlides) continue;
      const section = document.createElement("section");
      section.className = "orim-pptx-slide";
      section.setAttribute("aria-label", `Slide ${slide.number}`);
      section.dataset.slide = String(slide.number);
      section.style.aspectRatio = `${presentation.width} / ${presentation.height}`;
      section.style.setProperty("--orim-pptx-width", String(presentation.width));
      const background = fillBackground(slide.background, context);
      if (background) section.style.background = background;
      for (const item of slide.elements) section.append(renderElement(item, frame, context));
      sections.set(slide.number, section);
      root.append(section);
    }
    root.addEventListener("click", (event) => {
      const link = (event.target as Element | null)?.closest?.("a[href^='#slide-']");
      if (!link) return;
      event.preventDefault();
      sections.get(Number(link.getAttribute("href")?.slice(7)))?.scrollIntoView();
    });
    container.replaceChildren(root);
    return { presentation, destroy: () => lease.cancel() };
  } catch (error) {
    lease.cancel();
    throw error;
  }
};

export const renderPptx = (
  input: PackageInput,
  container: HTMLElement,
  options: RenderPptxOptions = {},
): Promise<RenderedPresentation> =>
  renderParsed(
    () => parsePptx(input),
    container,
    options.signal,
    (presentation) =>
      renderPresentation(presentation, container, { hiddenSlides: options.hiddenSlides }),
  );
