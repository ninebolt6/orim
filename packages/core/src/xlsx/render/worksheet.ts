import { safeHref, validColor } from "../../common/sanitize";
import { hasValue, type Cell, type CellStyle, type CellValue, type XlsxWorkbook } from "../model";
import { columnName } from "../ref";
import { computeLayout, HEADER_WIDTH, layoutKey, visibleSpan } from "../layout";
import { formatCellValue } from "../format";
import { defaultCellStyle } from "../parse/styles";
import { acquireContainerLease } from "../../common/lease";

const MIN_COLUMN_WIDTH = 8;

export interface RenderedWorksheet {
  readonly workbook: XlsxWorkbook;
  readonly sheet: number;
  destroy(): void;
}

const borderStyle = (line: string): { width: number; style: string } => ({
  width: line === "thick" || line === "double" ? 3 : line.startsWith("medium") ? 2 : 1,
  style:
    line === "double"
      ? "double"
      : line === "dotted"
        ? "dotted"
        : line.includes("ash")
          ? "dashed"
          : "solid",
});

const fontFamily = (name: string | undefined): string | undefined => {
  if (!name) return undefined;
  const clean = name.replace(/[\r\n"\\]/g, "");
  return clean ? `"${clean}", sans-serif` : undefined;
};

const fontStyle = (element: HTMLElement, font: Partial<CellStyle["font"]>): void => {
  const family = fontFamily(font.name);
  if (family) element.style.fontFamily = family;
  if (font.size !== undefined && Number.isFinite(font.size) && font.size > 0)
    element.style.fontSize = `${font.size}pt`;
  if (font.bold) element.style.fontWeight = "bold";
  if (font.italic) element.style.fontStyle = "italic";
  const decorations = [font.underline ? "underline" : "", font.strike ? "line-through" : ""].filter(
    Boolean,
  );
  if (decorations.length) {
    element.style.textDecorationLine = decorations.join(" ");
    if (font.underline === "double") element.style.textDecorationStyle = "double";
  }
  if (font.vertAlign)
    element.style.verticalAlign = font.vertAlign === "superscript" ? "super" : "sub";
  const color = validColor(font.color);
  if (color) element.style.color = color;
};

const setBorder = (
  cell: HTMLTableCellElement,
  side: "top" | "right" | "bottom" | "left",
  border: CellStyle["border"][keyof CellStyle["border"]],
): void => {
  if (!border) return;
  const style = borderStyle(border.line);
  const color = validColor(border.color) ?? "#000000";
  cell.style.setProperty(`border-${side}`, `${style.width}px ${style.style} ${color}`);
};

const renderText = (
  container: HTMLElement,
  cell: Cell | undefined,
  value: CellValue,
  formatted: ReturnType<typeof formatCellValue>,
): void => {
  if (cell?.runs && value.type === "string") {
    for (const run of cell.runs) {
      const span = container.ownerDocument.createElement("span");
      fontStyle(span, run.font ?? {});
      span.append(container.ownerDocument.createTextNode(run.text));
      container.append(span);
    }
    return;
  }
  container.append(container.ownerDocument.createTextNode(formatted.text));
  const color = validColor(formatted.color);
  if (color) container.style.color = color;
};

const effectiveAlignment = (
  style: CellStyle,
  value: CellValue,
): "left" | "center" | "right" | "justify" => {
  const alignment = style.alignment.horizontal;
  if (alignment === "fill") return "left";
  if (alignment === "centerContinuous") return "center";
  if (alignment === "distributed" || alignment === "justify") return "justify";
  if (alignment === "left" || alignment === "center" || alignment === "right") return alignment;
  if (value.type === "number") return "right";
  if (value.type === "boolean" || value.type === "error") return "center";
  return "left";
};

const applyStickyZIndex = (
  element: HTMLElement,
  isHeader: boolean,
  stickyTop: boolean,
  stickyLeft: boolean,
): void => {
  if (!stickyTop && !stickyLeft) return;
  element.style.zIndex = String((isHeader ? 4 : 0) + (stickyTop ? 2 : 0) + (stickyLeft ? 1 : 0));
};

export const renderWorksheet = (
  workbook: XlsxWorkbook,
  sheetIndex: number,
  container: HTMLElement,
): RenderedWorksheet => {
  const worksheet = workbook.sheets[sheetIndex];
  if (!worksheet) throw new RangeError("Worksheet index is out of range.");
  const layout = computeLayout(worksheet);
  const truncated = layout.truncated;
  const worksheetRows = new Map(worksheet.rows.map((row) => [row.index, row]));
  const worksheetCells = new Map<string, Cell>();
  for (const row of worksheet.rows)
    for (const cell of row.cells) worksheetCells.set(layoutKey(row.index, cell.column), cell);
  let root: HTMLDivElement | undefined;
  const lease = acquireContainerLease(container, () => {
    if (root?.parentElement === container) container.replaceChildren();
  });
  try {
    const document = container.ownerDocument;
    root = document.createElement("div");
    root.className = "orim-xlsx";
    const scroll = document.createElement("div");
    scroll.className = "orim-xlsx-scroll";
    const table = document.createElement("table");
    table.className = "orim-xlsx-grid";
    table.style.tableLayout = "fixed";
    table.style.borderCollapse = "collapse";
    const showHeaders = worksheet.showHeaders;
    const headerWidth = showHeaders ? HEADER_WIDTH : 0;
    const widths = layout.columns.map((column) => column.width);
    const columnElements: HTMLTableColElement[] = [];
    const resizeHandles: HTMLDivElement[] = [];
    table.style.width = `${headerWidth + widths.reduce((total, width) => total + width, 0)}px`;
    const colgroup = document.createElement("colgroup");
    if (showHeaders) {
      const rowHeaderColumn = document.createElement("col");
      rowHeaderColumn.style.width = `${HEADER_WIDTH}px`;
      colgroup.append(rowHeaderColumn);
    }
    for (const [columnIndex] of layout.columns.entries()) {
      const col = document.createElement("col");
      col.style.width = `${widths[columnIndex]}px`;
      colgroup.append(col);
      columnElements.push(col);
    }
    table.append(colgroup);
    const stickyCellsByColumn = layout.columns.map((): HTMLElement[] => []);
    type BoundedSpill = { content: HTMLElement; from: number; to: number; inset: number };
    const boundedSpillsByColumn = layout.columns.map((): BoundedSpill[] => []);
    const spillWidth = (from: number, to: number, inset: number): number => {
      let width = -inset;
      for (let index = from; index <= to; index++) width += widths[index]!;
      return Math.max(0, width);
    };
    const resizeColumn = (columnIndex: number, width: number): void => {
      const resizedWidth = Math.max(MIN_COLUMN_WIDTH, width);
      widths[columnIndex] = resizedWidth;
      columnElements[columnIndex]!.style.width = `${resizedWidth}px`;
      resizeHandles[columnIndex]?.setAttribute("aria-valuenow", String(Math.round(resizedWidth)));
      table.style.width = `${headerWidth + widths.reduce((total, current) => total + current, 0)}px`;
      for (const spill of boundedSpillsByColumn[columnIndex]!) {
        spill.content.style.maxWidth = `${spillWidth(spill.from, spill.to, spill.inset)}px`;
      }
      if (layout.columns[columnIndex]!.stickyLeft === undefined) return;
      let left = headerWidth;
      for (let stickyIndex = 0; stickyIndex < layout.columns.length; stickyIndex++) {
        if (layout.columns[stickyIndex]!.stickyLeft === undefined) break;
        if (stickyIndex > columnIndex) {
          for (const element of stickyCellsByColumn[stickyIndex]!) element.style.left = `${left}px`;
        }
        left += widths[stickyIndex]!;
      }
    };
    if (showHeaders) {
      const head = document.createElement("thead");
      const row = document.createElement("tr");
      row.className = "orim-xlsx-header-row";
      const corner = document.createElement("th");
      corner.className = "orim-xlsx-corner";
      corner.scope = "col";
      corner.style.width = `${HEADER_WIDTH}px`;
      applyStickyZIndex(corner, true, true, true);
      row.append(corner);
      for (const [columnIndex, column] of layout.columns.entries()) {
        const header = document.createElement("th");
        header.className = "orim-xlsx-column-header";
        header.scope = "col";
        header.append(document.createTextNode(columnName(column.index)));
        if (column.stickyLeft !== undefined) {
          header.style.left = `${column.stickyLeft}px`;
          stickyCellsByColumn[columnIndex]!.push(header);
        }
        applyStickyZIndex(header, true, true, column.stickyLeft !== undefined);
        const handle = document.createElement("div");
        handle.className = "orim-xlsx-resize";
        handle.setAttribute("role", "separator");
        handle.setAttribute("aria-orientation", "vertical");
        handle.setAttribute("aria-label", `Resize column ${columnName(column.index)}`);
        handle.tabIndex = 0;
        handle.setAttribute(
          "aria-valuemin",
          String(Math.min(MIN_COLUMN_WIDTH, Math.round(widths[columnIndex]!))),
        );
        handle.setAttribute("aria-valuenow", String(Math.round(widths[columnIndex]!)));
        resizeHandles[columnIndex] = handle;
        handle.addEventListener("keydown", (event: KeyboardEvent) => {
          if (event.key === "ArrowRight") resizeColumn(columnIndex, widths[columnIndex]! + 8);
          else if (event.key === "ArrowLeft") resizeColumn(columnIndex, widths[columnIndex]! - 8);
          else return;
          event.preventDefault();
        });
        let activePointer: number | undefined;
        let startX = 0;
        let startWidth = 0;
        handle.addEventListener("pointerdown", (event: PointerEvent) => {
          event.preventDefault();
          activePointer = event.pointerId;
          startX = event.clientX;
          startWidth = widths[columnIndex]!;
          handle.setPointerCapture(event.pointerId);
        });
        handle.addEventListener("pointermove", (event: PointerEvent) => {
          if (activePointer !== event.pointerId) return;
          resizeColumn(
            columnIndex,
            Math.max(MIN_COLUMN_WIDTH, startWidth + event.clientX - startX),
          );
        });
        const stopResizing = (event: PointerEvent): void => {
          if (activePointer !== event.pointerId) return;
          if (handle.hasPointerCapture(event.pointerId))
            handle.releasePointerCapture(event.pointerId);
          activePointer = undefined;
        };
        handle.addEventListener("pointerup", stopResizing);
        handle.addEventListener("pointercancel", stopResizing);
        header.append(handle);
        row.append(header);
      }
      head.append(row);
      table.append(head);
    }
    const body = document.createElement("tbody");
    const hyperlinks = new Map<string, string>();
    const rowHeightPrefix: number[] = [0];
    for (const row of layout.rows)
      rowHeightPrefix.push(rowHeightPrefix[rowHeightPrefix.length - 1]! + row.height);
    for (const link of worksheet.hyperlinks) {
      const href = safeHref(link.href);
      if (!href) continue;
      const [rowFrom, rowTo] = visibleSpan(layout.rows, link.range.top, link.range.bottom);
      const [columnFrom, columnTo] = visibleSpan(layout.columns, link.range.left, link.range.right);
      for (const row of layout.rows.slice(rowFrom, rowTo)) {
        for (const column of layout.columns.slice(columnFrom, columnTo)) {
          const key = layoutKey(row.index, column.index);
          if (!layout.covered.has(key) && !hyperlinks.has(key)) hyperlinks.set(key, href);
        }
      }
    }
    for (const rowLayout of layout.rows) {
      const rowElement = document.createElement("tr");
      rowElement.className = "orim-xlsx-row";
      rowElement.style.height = `${rowLayout.height}pt`;
      if (showHeaders) {
        const header = document.createElement("th");
        header.className = "orim-xlsx-row-header";
        header.scope = "row";
        header.append(document.createTextNode(String(rowLayout.index + 1)));
        header.style.width = `${HEADER_WIDTH}px`;
        header.style.position = "sticky";
        header.style.left = "0";
        if (rowLayout.stickyTop !== undefined) {
          header.style.top = `${rowLayout.stickyTop}px`;
        }
        applyStickyZIndex(header, true, rowLayout.stickyTop !== undefined, true);
        rowElement.append(header);
      }
      for (let columnIndex = 0; columnIndex < layout.columns.length; columnIndex++) {
        const columnLayout = layout.columns[columnIndex]!;
        const key = layoutKey(rowLayout.index, columnLayout.index);
        if (layout.covered.has(key)) continue;
        const merge = layout.mergeAt.get(key);
        const cell = merge?.cell ?? worksheetCells.get(key);
        const cellRow = merge
          ? worksheetRows.get(merge.range.top)
          : worksheetRows.get(rowLayout.index);
        const columnStyle = columnLayout.style;
        const style = cell?.style ?? cellRow?.style ?? columnStyle ?? defaultCellStyle;
        const value = cell?.value ?? { type: "empty" as const };
        const alignment = effectiveAlignment(style, value);
        const formatted = formatCellValue(value, style.numberFormat, workbook.date1904);
        const td = document.createElement("td");
        td.className = "orim-xlsx-cell";
        if (merge) {
          td.rowSpan = merge.rowSpan;
          td.colSpan = merge.columnSpan;
        }
        const fill = validColor(style.fill);
        if (fill) td.style.backgroundColor = fill;
        if (worksheet.showGridLines && !fill) td.classList.add("orim-xlsx-gridline");
        if (style.border) {
          setBorder(td, "top", style.border.top);
          setBorder(td, "right", style.border.right);
          setBorder(td, "bottom", style.border.bottom);
          setBorder(td, "left", style.border.left);
        }
        fontStyle(td, style.font);
        const renderedColor = validColor(formatted.color);
        if (renderedColor) td.style.color = renderedColor;
        td.style.textAlign = alignment;
        td.style.verticalAlign = style.alignment.vertical ?? "bottom";
        td.style.whiteSpace = style.alignment.wrap ? "pre-wrap" : "pre";
        if (style.alignment.indent > 0) td.style.paddingLeft = `${style.alignment.indent * 9}px`;
        const content = document.createElement("div");
        content.className = "orim-xlsx-content";
        let spanHeight = rowLayout.height;
        if (merge) {
          const [from, to] = visibleSpan(layout.rows, merge.range.top, merge.range.bottom);
          spanHeight = rowHeightPrefix[to]! - rowHeightPrefix[from]!;
        }
        content.style.height = `${spanHeight}pt`;
        content.style.overflow = "visible clip";
        const hyperlink = hyperlinks.get(key);
        if (hyperlink) {
          const anchor = document.createElement("a");
          anchor.href = hyperlink;
          anchor.target = "_blank";
          anchor.rel = "noopener noreferrer";
          renderText(anchor, cell, value, formatted);
          content.append(anchor);
        } else {
          renderText(content, cell, value, formatted);
        }
        td.append(content);
        const mergeCrossesFreeze =
          merge !== undefined &&
          ((merge.range.top < worksheet.frozen.rows &&
            merge.range.bottom >= worksheet.frozen.rows) ||
            (merge.range.left < worksheet.frozen.columns &&
              merge.range.right >= worksheet.frozen.columns));
        const rowSticky = !mergeCrossesFreeze
          ? layout.rowAt.get(rowLayout.index)?.stickyTop
          : undefined;
        const columnSticky = !mergeCrossesFreeze
          ? layout.columnAt.get(columnLayout.index)?.stickyLeft
          : undefined;
        if (rowSticky !== undefined || columnSticky !== undefined) {
          td.style.position = "sticky";
          if (rowSticky !== undefined) td.style.top = `${rowSticky}px`;
          if (columnSticky !== undefined) {
            td.style.left = `${columnSticky}px`;
            stickyCellsByColumn[columnIndex]!.push(td);
          }
          td.style.backgroundColor = fill ?? "#ffffff";
          applyStickyZIndex(td, false, rowSticky !== undefined, columnSticky !== undefined);
        }
        const canSpill =
          value.type === "string" &&
          value.value.length > 0 &&
          !style.alignment.wrap &&
          alignment === "left" &&
          !merge;
        if (canSpill) {
          let spillTo = columnIndex;
          let bounded = false;
          while (spillTo + 1 < layout.columns.length) {
            const nextColumn = layout.columns[spillTo + 1]!;
            const nextKey = layoutKey(rowLayout.index, nextColumn.index);
            if (
              layout.covered.has(nextKey) ||
              layout.mergeAt.has(nextKey) ||
              hasValue(worksheetCells.get(nextKey)?.value ?? { type: "empty" })
            ) {
              bounded = true;
              break;
            }
            spillTo += 1;
          }
          if (spillTo > columnIndex) {
            td.classList.add("orim-xlsx-spill");
            if (bounded) {
              const inset = style.alignment.indent * 9;
              content.style.maxWidth = `${spillWidth(columnIndex, spillTo, inset)}px`;
              content.style.minWidth = "0";
              content.style.overflowX = "hidden";
              const spill = { content, from: columnIndex, to: spillTo, inset };
              for (let index = columnIndex; index <= spillTo; index++) {
                boundedSpillsByColumn[index]!.push(spill);
              }
            }
          }
        }
        rowElement.append(td);
      }
      body.append(rowElement);
    }
    table.append(body);
    scroll.append(table);
    if (truncated) {
      const note = document.createElement("div");
      note.className = "orim-xlsx-note";
      note.append(
        document.createTextNode(
          `Showing the first ${layout.rows.length} rows and ${layout.columns.length} columns.`,
        ),
      );
      scroll.append(note);
    }
    root.append(scroll);
    container.replaceChildren(root);
    return {
      workbook,
      sheet: sheetIndex,
      destroy: () => lease.cancel(),
    };
  } catch (error) {
    lease.cancel();
    throw error;
  }
};
