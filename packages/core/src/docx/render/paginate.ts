import type { Block, DocxDocument, SectionProperties } from "../model";
import type { PendingField } from "./context";

export interface Page {
  wrapper: HTMLElement;
  page: HTMLElement;
  content: HTMLElement;
  readonly properties: SectionProperties;
  readonly number: number;
}

export const COMMENT_LANE_WIDTH_PT = 168;
type RenderBlock = (block: Block, pageNumber?: number, decoration?: boolean) => HTMLElement;

const addHeaderFooter = (
  doc: DocxDocument,
  properties: SectionProperties,
  page: HTMLElement,
  pageNumber: number,
  firstPage: boolean,
  renderBlock: RenderBlock,
): void => {
  const { margins: margin, width } = properties;
  const usableWidth = Math.max(1, width - margin.left - margin.right);
  const parityEven = pageNumber % 2 === 0;
  const header =
    properties.titlePage && firstPage
      ? (properties.firstHeader ?? properties.header)
      : doc.evenAndOddHeaders && parityEven
        ? (properties.evenHeader ?? properties.header)
        : properties.header;
  const footer =
    properties.titlePage && firstPage
      ? (properties.firstFooter ?? properties.footer)
      : doc.evenAndOddHeaders && parityEven
        ? (properties.evenFooter ?? properties.footer)
        : properties.footer;
  if (header?.length) {
    const element = document.createElement("div");
    element.className = "orim-header";
    element.style.top = `${margin.header}pt`;
    element.style.left = `${margin.left}pt`;
    element.style.width = `${usableWidth}pt`;
    for (const block of header) element.append(renderBlock(block, pageNumber, true));
    page.append(element);
  }
  if (footer?.length) {
    const element = document.createElement("div");
    element.className = "orim-footer";
    element.style.bottom = `${margin.footer}pt`;
    element.style.left = `${margin.left}pt`;
    element.style.width = `${usableWidth}pt`;
    for (const block of footer) element.append(renderBlock(block, pageNumber, true));
    page.append(element);
  }
};

const applyPageGeometry = (page: Page, properties: SectionProperties): Page => {
  const { width, height, margins: margin } = properties;
  const usableWidth = Math.max(1, width - margin.left - margin.right);
  const usableHeight = Math.max(1, height - margin.top - margin.bottom);
  page.wrapper.style.setProperty("--orim-page-width", `${width}pt`);
  page.wrapper.style.height = `${height}pt`;
  page.page.style.width = `${width}pt`;
  page.page.style.height = `${height}pt`;
  page.content.style.left = `${margin.left}pt`;
  page.content.style.top = `${margin.top}pt`;
  page.content.style.width = `${usableWidth}pt`;
  page.content.style.height = `${usableHeight}pt`;
  if (properties.linePitch)
    page.content.style.setProperty("--orim-pitch", `${properties.linePitch}pt`);
  else page.content.style.removeProperty("--orim-pitch");
  return { ...page, properties };
};

const createPage = (
  pagesRoot: HTMLElement,
  pages: Page[],
  doc: DocxDocument,
  properties: SectionProperties,
  includeComments: boolean,
  renderBlock: RenderBlock,
): Page => {
  const wrapper = document.createElement("div");
  wrapper.className = "orim-page-wrap";
  const pageElement = document.createElement("div");
  pageElement.className = "orim-page";
  const content = document.createElement("div");
  content.className = "orim-content";
  pageElement.append(content);
  const pageNumber = pages.length + 1;
  const firstPage = !pages.some((page) => page.properties === properties);
  addHeaderFooter(doc, properties, pageElement, pageNumber, firstPage, renderBlock);
  if (includeComments) {
    const comments = document.createElement("div");
    comments.className = "orim-page-comments";
    wrapper.append(comments);
  }
  wrapper.append(pageElement);
  const page = applyPageGeometry(
    { wrapper, page: pageElement, content, properties, number: pageNumber },
    properties,
  );
  pagesRoot.append(wrapper);
  pages.push(page);
  return page;
};

const overflows = (element: HTMLElement, page: Page): boolean => {
  const bounds = page.content.getBoundingClientRect();
  const blockBounds = element.getBoundingClientRect();
  return (
    page.content.scrollHeight > page.content.clientHeight + 1 ||
    blockBounds.bottom > bounds.bottom + 1
  );
};

const splitTable = (table: HTMLTableElement, page: Page, newPage: () => Page): Page => {
  if (!overflows(table, page)) return page;
  const originalRows = Array.from(table.rows);
  const headerRows = originalRows.filter((row) => row.dataset.orimHeader === "true");
  const movable = originalRows.filter((row) => row.dataset.orimHeader !== "true");
  if (!movable.length) return page;
  const rowParents = new Map(originalRows.map((row) => [row, row.parentElement]));
  const bounds = page.content.getBoundingClientRect();
  const rowBottoms = movable.map((row) => row.getBoundingClientRect().bottom);
  const firstOverflowing = rowBottoms.findIndex((bottom) => bottom > bounds.bottom + 1);
  let keep = Math.max(1, firstOverflowing < 0 ? movable.length : firstOverflowing);
  const continuedRows = movable.slice(keep);
  for (const row of continuedRows) row.remove();
  while (overflows(table, page) && keep > 1) {
    const row = movable[--keep]!;
    row.remove();
    continuedRows.unshift(row);
  }
  while (continuedRows.length) {
    const row = continuedRows[0]!;
    const parent = rowParents.get(row);
    if (!parent) break;
    parent.append(row);
    if (overflows(table, page)) {
      row.remove();
      break;
    }
    continuedRows.shift();
    keep++;
  }
  if (overflows(table, page) && page.content.childElementCount > 1) {
    for (const row of originalRows) {
      const parent = rowParents.get(row);
      if (parent && row.parentElement !== parent) parent.append(row);
    }
    table.remove();
    const next = newPage();
    next.content.append(table);
    return splitTable(table, next, newPage);
  }
  if (!continuedRows.length) return page;
  const next = newPage();
  const continuation = table.cloneNode(false) as HTMLTableElement;
  for (const child of Array.from(table.children))
    if (child.localName === "colgroup") continuation.append(child.cloneNode(true));
  const body = document.createElement("tbody");
  for (const row of headerRows) body.append(row.cloneNode(true));
  for (const row of continuedRows) body.append(row);
  continuation.append(body);
  next.content.append(continuation);
  return splitTable(continuation, next, newPage);
};

export const paginate = (
  root: HTMLElement,
  doc: DocxDocument,
  mode: string,
  renderBlock: (block: Block, pageNumber?: number, decoration?: boolean) => HTMLElement,
  includeComments: boolean,
  fields: readonly PendingField[],
): Page[] => {
  root.replaceChildren();
  root.className = `orim-root orim-${mode}`;
  root.style.setProperty(
    "--orim-comment-lane",
    includeComments ? `${COMMENT_LANE_WIDTH_PT}pt` : "0pt",
  );
  const pagesRoot = document.createElement("div");
  pagesRoot.className = "orim-pages";
  root.append(pagesRoot);
  const pages: Page[] = [];

  for (let sectionIndex = 0; sectionIndex < doc.sections.length; sectionIndex++) {
    const section = doc.sections[sectionIndex];
    const props = section.properties;
    const newPage = () => createPage(pagesRoot, pages, doc, props, includeComments, renderBlock);

    const breakType = props.breakType ?? "nextPage";
    let current: Page;
    if (sectionIndex > 0 && breakType === "continuous" && pages.length) {
      const previous = pages[pages.length - 1];
      current = applyPageGeometry(previous, props);
      pages[pages.length - 1] = current;
      current.page
        .querySelectorAll(".orim-header,.orim-footer")
        .forEach((element) => element.remove());
      const firstPage = !pages.some((item) => item !== current && item.properties === props);
      addHeaderFooter(doc, props, current.page, current.number, firstPage, renderBlock);
    } else {
      if (sectionIndex > 0 && breakType === "oddPage" && pages.length % 2 === 1) newPage();
      if (sectionIndex > 0 && breakType === "evenPage" && pages.length % 2 === 0) newPage();
      current = newPage();
    }
    for (let blockIndex = 0; blockIndex < section.blocks.length; blockIndex++) {
      const block = section.blocks[blockIndex];
      if (
        block.type === "paragraph" &&
        block.properties.pageBreakBefore &&
        current.content.childElementCount
      )
        current = newPage();
      const element = renderBlock(block, current.number);
      current.content.append(element);
      if (element instanceof HTMLTableElement) {
        current = splitTable(element, current, newPage);
      } else if (overflows(element, current) && current.content.childElementCount > 1) {
        element.remove();
        current = newPage();
        current.content.append(element);
      }
      if (
        block.type === "paragraph" &&
        block.properties.pageBreakAfter &&
        blockIndex < section.blocks.length - 1
      )
        current = newPage();
    }
  }

  const total = pages.length;
  for (const field of fields) {
    if (field.name === "NUMPAGES") field.element.textContent = String(total);
    else if (field.name === "PAGE") {
      const page = pages.find((item) => item.content.contains(field.element));
      field.element.textContent = page ? String(page.number) : field.cached;
    }
  }
  return pages;
};
