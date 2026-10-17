import { safeHref, validColor } from "../../common/sanitize";
import type { Cell, CellStyle, CellValue, XlsxWorkbook } from "../model";
import { columnName } from "../ref";
import { computeLayout, HEADER_WIDTH, layoutKey } from "../layout";
import { formatCellValue } from "../format";
import { defaultCellStyle } from "../parse/styles";
import { acquireContainerLease } from "../../common/lease";

export interface RenderedWorksheet {
  readonly workbook: XlsxWorkbook;
  readonly sheet: number;
  destroy(): void;
}

const borderStyles: Readonly<Record<string, { width: number; style: string }>> = {
  thin: { width: 1, style: "solid" },
  hair: { width: 1, style: "solid" },
  medium: { width: 2, style: "solid" },
  thick: { width: 3, style: "solid" },
  dashed: { width: 1, style: "dashed" },
  dashDot: { width: 1, style: "dashed" },
  dashDotDot: { width: 1, style: "dashed" },
  slantDashDot: { width: 1, style: "dashed" },
  mediumDashed: { width: 2, style: "dashed" },
  mediumDashDot: { width: 2, style: "dashed" },
  mediumDashDotDot: { width: 2, style: "dashed" },
  dotted: { width: 1, style: "dotted" },
  double: { width: 3, style: "double" },
};

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
  const style = borderStyles[border.line];
  if (!style) return;
  const color = validColor(border.color) ?? "#000000";
  cell.style.setProperty(`border-${side}`, `${style.width}px ${style.style} ${color}`);
};

const visibleValue = (value: CellValue): boolean => {
  return value.type !== "empty" && (value.type !== "string" || value.value.length > 0);
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

const MAX_RENDERED_CELLS = 20_000;

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
  const visibleRows =
    layout.columns.length > 0
      ? layout.rows.slice(0, Math.max(1, Math.floor(MAX_RENDERED_CELLS / layout.columns.length)))
      : layout.rows;
  const truncated = layout.truncated || visibleRows.length < layout.rows.length;
  const worksheetRows = new Map(worksheet.rows.map((row) => [row.index, row]));
  const worksheetCells = new Map<string, Cell>();
  for (const row of worksheet.rows)
    for (const cell of row.cells) worksheetCells.set(layoutKey(row.index, cell.column), cell);
  const renderedMergeRows = new Map<string, (typeof layout.rows)[number][]>();
  for (const merge of layout.merges) {
    const rows = visibleRows.filter(
      (row) => row.index >= merge.range.top && row.index <= merge.range.bottom,
    );
    if (rows.length) renderedMergeRows.set(layoutKey(merge.anchorRow, merge.anchorColumn), rows);
  }
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
    table.style.width = `${headerWidth + layout.columns.reduce((total, column) => total + column.width, 0)}px`;
    const colgroup = document.createElement("colgroup");
    if (showHeaders) {
      const rowHeaderColumn = document.createElement("col");
      rowHeaderColumn.style.width = `${HEADER_WIDTH}px`;
      colgroup.append(rowHeaderColumn);
    }
    for (const column of layout.columns) {
      const col = document.createElement("col");
      col.style.width = `${column.width}px`;
      colgroup.append(col);
    }
    table.append(colgroup);
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
      for (const column of layout.columns) {
        const header = document.createElement("th");
        header.className = "orim-xlsx-column-header";
        header.scope = "col";
        header.append(document.createTextNode(columnName(column.index)));
        if (column.stickyLeft !== undefined) header.style.left = `${column.stickyLeft}px`;
        applyStickyZIndex(header, true, true, column.stickyLeft !== undefined);
        row.append(header);
      }
      head.append(row);
      table.append(head);
    }
    const body = document.createElement("tbody");
    const hyperlinks = new Map<string, string>();
    for (const link of worksheet.hyperlinks) {
      const href = safeHref(link.href);
      if (!href) continue;
      for (const row of visibleRows) {
        if (row.index < link.range.top || row.index > link.range.bottom) continue;
        for (const column of layout.columns) {
          if (column.index < link.range.left || column.index > link.range.right) continue;
          const key = layoutKey(row.index, column.index);
          if (!layout.covered.has(key) && !hyperlinks.has(key)) hyperlinks.set(key, href);
        }
      }
    }
    for (const rowLayout of visibleRows) {
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
        const mergeRows = merge ? renderedMergeRows.get(key) : undefined;
        const cell = merge?.cell ?? worksheetCells.get(key);
        const cellRow = merge
          ? worksheetRows.get(merge.range.top)
          : worksheetRows.get(rowLayout.index);
        const columnStyle = columnLayout.style;
        const style = cell?.style ?? cellRow?.style ?? columnStyle ?? defaultCellStyle;
        const value = cell?.value ?? { type: "empty" as const };
        const formatted = formatCellValue(value, style.numberFormat, workbook.date1904);
        const td = document.createElement("td");
        td.className = "orim-xlsx-cell";
        if (merge && mergeRows) {
          td.rowSpan = mergeRows.length;
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
        td.style.textAlign = effectiveAlignment(style, value);
        td.style.verticalAlign = style.alignment.vertical ?? "bottom";
        td.style.whiteSpace = style.alignment.wrap ? "pre-wrap" : "pre";
        if (style.alignment.indent > 0) td.style.paddingLeft = `${style.alignment.indent * 9}px`;
        const content = document.createElement("div");
        content.className = "orim-xlsx-content";
        const spanHeight = mergeRows
          ? mergeRows.reduce((total, candidate) => total + candidate.height, 0)
          : rowLayout.height;
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
          if (columnSticky !== undefined) td.style.left = `${columnSticky}px`;
          td.style.backgroundColor = fill ?? "#ffffff";
          applyStickyZIndex(td, false, rowSticky !== undefined, columnSticky !== undefined);
        }
        const nextColumn = layout.columns[columnIndex + 1];
        const spill =
          value.type === "string" &&
          value.value.length > 0 &&
          !style.alignment.wrap &&
          effectiveAlignment(style, value) === "left" &&
          !merge &&
          nextColumn !== undefined &&
          !layout.covered.has(layoutKey(rowLayout.index, nextColumn.index)) &&
          !layout.mergeAt.has(layoutKey(rowLayout.index, nextColumn.index)) &&
          !visibleValue(
            worksheetCells.get(layoutKey(rowLayout.index, nextColumn.index))?.value ?? {
              type: "empty",
            },
          );
        if (spill) td.classList.add("orim-xlsx-spill");
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
          `Showing the first ${visibleRows.length} rows and ${layout.columns.length} columns.`,
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
