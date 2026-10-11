import type { Block, Table } from "../docx/model";
import type { RenderContext } from "./context";
import { validColor } from "./sanitize";
import { changeColor } from "./text-style";

type RenderBlock = (block: Block, context: RenderContext) => HTMLElement;

export function renderTable(
  block: Table,
  context: RenderContext,
  renderBlock: RenderBlock,
): HTMLTableElement {
  const table = document.createElement("table");
  table.className = "orim-table";
  if (block.width) table.style.width = block.width;
  if (block.align === "center") table.style.marginInline = "auto";
  else if (block.align === "right") table.style.marginLeft = "auto";
  else if (block.indent) table.style.marginLeft = `${block.indent}pt`;
  const colgroup = document.createElement("colgroup");
  for (const width of block.widths) {
    const col = document.createElement("col");
    if (width) col.style.width = width;
    colgroup.append(col);
  }
  if (block.widths.length) table.append(colgroup);
  const body = document.createElement("tbody");
  for (const [rowIndex, row] of block.rows.entries()) {
    const tr = document.createElement("tr");
    if (row.header) tr.dataset.orimHeader = "true";
    if (row.height) tr.style.height = `${row.height}pt`;
    if (row.change) {
      tr.classList.add("orim-table-row--change", `orim-table-row--${row.change.kind}`);
      tr.style.setProperty("--orim-change-color", changeColor(row.change, context.colors));
      tr.title = `${row.change.kind === "ins" ? "Inserted" : "Deleted"} by ${row.change.author}`;
    }
    for (const [cellIndex, cell] of row.cells.entries()) {
      if (cell.rowSpan === 0) continue;
      const td = document.createElement("td");
      td.className = "orim-table-cell";
      if (cell.gridSpan > 1) td.colSpan = cell.gridSpan;
      if (cell.rowSpan > 1) td.rowSpan = cell.rowSpan;
      if (cell.width) td.style.width = cell.width;
      const bg = validColor(cell.shading);
      if (bg) td.style.backgroundColor = bg;
      if (cell.verticalAlign) td.style.verticalAlign = cell.verticalAlign;
      for (const side of ["top", "right", "bottom", "left"] as const) {
        const tableSide =
          side === "top"
            ? rowIndex === 0
              ? "top"
              : "insideH"
            : side === "bottom"
              ? rowIndex === block.rows.length - 1
                ? "bottom"
                : "insideH"
              : side === "left"
                ? cellIndex === 0
                  ? "left"
                  : "insideV"
                : cellIndex === row.cells.length - 1
                  ? "right"
                  : "insideV";
        const border = cell.borders?.[side] ?? block.borders?.[tableSide];
        if (border) td.style.setProperty(`border-${side}`, border);
        const margin =
          cell.margins?.[side] ??
          block.margins?.[side] ??
          (side === "left" || side === "right" ? 5.4 : 0);
        td.style.setProperty(`padding-${side}`, `${margin}pt`);
      }
      for (const nested of cell.blocks) td.append(renderBlock(nested, context));
      tr.append(td);
    }
    body.append(tr);
  }
  table.append(body);
  return table;
}
