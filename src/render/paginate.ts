import type { Block, DocxDocument, SectionProperties } from "../docx/model";

export interface Page {
  wrapper: HTMLElement;
  page: HTMLElement;
  content: HTMLElement;
  properties: SectionProperties;
  number: number;
}

export function paginate(
  root: HTMLElement,
  doc: DocxDocument,
  mode: string,
  renderBlock: (block: Block) => HTMLElement,
  includeComments: boolean,
): Page[] {
  root.replaceChildren();
  root.className = `orim-root orim-${mode}`;
  const pagesRoot = document.createElement("div");
  pagesRoot.className = "orim-pages";
  root.append(pagesRoot);
  const pages: Page[] = [];

  for (let sectionIndex = 0; sectionIndex < doc.sections.length; sectionIndex++) {
    const section = doc.sections[sectionIndex];
    const props = section.properties;
    const width = props.width;
    const height = props.height;
    const margin = props.margins;
    const usableWidth = Math.max(1, width - margin.left - margin.right);
    const usableHeight = Math.max(1, height - margin.top - margin.bottom);

    const addHeaderFooter = (page: HTMLElement, pageNumber: number, firstPage: boolean): void => {
      const parityEven = pageNumber % 2 === 0;
      const header =
        props.titlePage && firstPage
          ? (props.firstHeader ?? props.header)
          : doc.evenAndOddHeaders && parityEven
            ? (props.evenHeader ?? props.header)
            : props.header;
      const footer =
        props.titlePage && firstPage
          ? (props.firstFooter ?? props.footer)
          : doc.evenAndOddHeaders && parityEven
            ? (props.evenFooter ?? props.footer)
            : props.footer;
      if (header?.length) {
        const el = document.createElement("div");
        el.className = "orim-header";
        el.style.top = `${margin.header}pt`;
        el.style.left = `${margin.left}pt`;
        el.style.width = `${usableWidth}pt`;
        for (const block of header) el.append(renderBlock(block));
        page.append(el);
      }
      if (footer?.length) {
        const el = document.createElement("div");
        el.className = "orim-footer";
        el.style.bottom = `${margin.footer}pt`;
        el.style.left = `${margin.left}pt`;
        el.style.width = `${usableWidth}pt`;
        for (const block of footer) el.append(renderBlock(block));
        page.append(el);
      }
    };

    const newPage = (): Page => {
      const wrapper = document.createElement("div");
      wrapper.className = "orim-page-wrap";
      wrapper.style.width = `${width + (includeComments ? 224 : 0)}pt`;
      wrapper.style.height = `${height}pt`;
      const page = document.createElement("div");
      page.className = "orim-page";
      page.style.width = `${width}pt`;
      page.style.height = `${height}pt`;
      const content = document.createElement("div");
      content.className = "orim-content";
      content.style.left = `${margin.left}pt`;
      content.style.top = `${margin.top}pt`;
      content.style.width = `${usableWidth}pt`;
      content.style.height = `${usableHeight}pt`;
      page.append(content);

      const pageNumber = pages.length + 1;
      const firstPage = !pages.some((item) => item.properties === props);
      addHeaderFooter(page, pageNumber, firstPage);
      if (includeComments) {
        const comments = document.createElement("div");
        comments.className = "orim-page-comments";
        wrapper.append(comments);
      }
      wrapper.append(page);
      pagesRoot.append(wrapper);
      const result = { wrapper, page, content, properties: props, number: pageNumber };
      pages.push(result);
      return result;
    };

    const previousBreak = doc.sections[sectionIndex - 1]?.properties.breakType ?? "nextPage";
    let current: Page;
    if (sectionIndex > 0 && previousBreak === "continuous" && pages.length) {
      current = pages[pages.length - 1];
      current.properties = props;
      current.wrapper.style.width = `${width + (includeComments ? 224 : 0)}pt`;
      current.wrapper.style.height = `${height}pt`;
      current.page.style.width = `${width}pt`;
      current.page.style.height = `${height}pt`;
      current.content.style.left = `${margin.left}pt`;
      current.content.style.top = `${margin.top}pt`;
      current.content.style.width = `${usableWidth}pt`;
      current.content.style.height = `${usableHeight}pt`;
      current.page
        .querySelectorAll(".orim-header,.orim-footer")
        .forEach((element) => element.remove());
      const firstPage = !pages.some((item) => item !== current && item.properties === props);
      addHeaderFooter(current.page, current.number, firstPage);
    } else {
      if (sectionIndex > 0 && previousBreak === "oddPage" && pages.length % 2 === 1) newPage();
      if (sectionIndex > 0 && previousBreak === "evenPage" && pages.length % 2 === 0) newPage();
      current = newPage();
    }
    const overflows = (element: HTMLElement, page: Page): boolean => {
      const bounds = page.content.getBoundingClientRect();
      const blockBounds = element.getBoundingClientRect();
      return (
        page.content.scrollHeight > page.content.clientHeight + 1 ||
        blockBounds.bottom > bounds.bottom + 1
      );
    };
    const splitTable = (table: HTMLTableElement, page: Page): void => {
      if (!overflows(table, page) || table.rows.length < 2) return;
      const headerRows = Array.from(table.rows).filter((row) => row.dataset.orimHeader === "true");
      const movable = Array.from(table.rows).filter((row) => row.dataset.orimHeader !== "true");
      if (!movable.length) return;
      const continuedRows: HTMLTableRowElement[] = [];
      while (overflows(table, page) && movable.length > 1) {
        const row = movable.pop()!;
        row.remove();
        continuedRows.unshift(row);
      }
      if (!continuedRows.length) return;
      const next = newPage();
      const continuation = table.cloneNode(false) as HTMLTableElement;
      for (const child of Array.from(table.children)) {
        if (child.localName === "colgroup") continuation.append(child.cloneNode(true));
      }
      const body = document.createElement("tbody");
      for (const row of headerRows) body.append(row.cloneNode(true));
      for (const row of continuedRows) body.append(row);
      continuation.append(body);
      next.content.append(continuation);
      splitTable(continuation, next);
    };

    for (let blockIndex = 0; blockIndex < section.blocks.length; blockIndex++) {
      const block = section.blocks[blockIndex];
      const pageBreaks =
        block.type === "paragraph" && block.runs.some((run) => run.type === "break" && run.page);
      const onlyBreak =
        block.type === "paragraph" &&
        pageBreaks &&
        block.runs.every(
          (run) => (run.type === "break" && run.page) || (run.type === "text" && !run.text.trim()),
        );
      if (
        block.type === "paragraph" &&
        block.properties.pageBreakBefore &&
        current.content.childElementCount
      )
        current = newPage();
      if (onlyBreak) {
        if (current.content.childElementCount) current = newPage();
        continue;
      }
      if (pageBreaks && current.content.childElementCount) current = newPage();
      const element = renderBlock(block);
      current.content.append(element);
      if (overflows(element, current) && current.content.childElementCount > 1) {
        element.remove();
        current = newPage();
        current.content.append(element);
      }
      if (element instanceof HTMLTableElement) splitTable(element, current);
      if (pageBreaks && blockIndex < section.blocks.length - 1) current = newPage();
    }
  }

  const total = pages.length;
  for (const element of root.querySelectorAll<HTMLElement>("[data-field]")) {
    const name = element.dataset.field;
    if (name === "PAGE")
      element.textContent = element.closest(".orim-page-wrap")?.querySelector(".orim-page")
        ? String(
            pages.find((page) => page.wrapper === element.closest(".orim-page-wrap"))?.number ?? "",
          )
        : "";
    else if (name === "NUMPAGES") element.textContent = String(total);
    else element.textContent = element.dataset.cached ?? "";
  }
  return pages;
}
