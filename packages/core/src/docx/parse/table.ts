import { attr, children, elements, first } from "../../xml";
import type { Block, Table, TableCell, TableRow } from "../model";
import type { ParseContext } from "./context";
import { border } from "../borders";
import { cssWidth, twips } from "../units";
import { applyContextualSpacing, blockNodes, splitAtPageBreaks } from "./blocks";
import { changeFrom } from "./inline";
import { parseParagraph } from "./paragraph";

function color(value: string | undefined): string | undefined {
  return value && /^[\da-f]{6}$/i.test(value) ? `#${value}` : undefined;
}

export function parseTable(node: Element, context: ParseContext, comments: Set<string>): Table {
  const tblPr = first(node, "tblPr");
  const styleId = attr(first(tblPr ?? node, "tblStyle"), "val");
  const style = context.styles.table(styleId);
  const grid = first(node, "tblGrid");
  const widths = elements(grid ?? node, "gridCol").map((col) => cssWidth(col));
  const directMargins: NonNullable<Table["margins"]> = {};
  const directTableMargins = first(tblPr ?? node, "tblCellMar");
  for (const side of ["top", "right", "bottom", "left"] as const) {
    const item = first(directTableMargins ?? node, side);
    const value = item ? twips(attr(item, "w"), 0) : undefined;
    if (value !== undefined) directMargins[side] = value;
  }
  const tableMargins = { ...style.margins, ...directMargins };
  const rows: TableRow[] = elements(node, "tr").map((row) => {
    const trPr = first(row, "trPr");
    const cells: TableCell[] = elements(row, "tc").map((cell) => {
      const tcPr = first(cell, "tcPr");
      const span = Number(attr(first(tcPr ?? cell, "gridSpan"), "val") ?? 1);
      const merge = first(tcPr ?? cell, "vMerge");
      const val = attr(merge, "val");
      const shading = attr(first(tcPr ?? cell, "shd"), "fill");
      const tcW = first(tcPr ?? cell, "tcW");
      const cellMargins = first(tcPr ?? cell, "tcMar");
      const borders: NonNullable<TableCell["borders"]> = {};
      const cellBorders = first(tcPr ?? cell, "tcBorders");
      for (const side of ["top", "right", "bottom", "left"] as const) {
        const value = border(first(cellBorders ?? cell, side));
        if (value !== undefined) borders[side] = value;
      }
      const blocks: Block[] = [];
      for (const child of blockNodes(children(cell))) {
        if (child.localName === "p")
          blocks.push(...splitAtPageBreaks(parseParagraph(child, context, comments)));
        else if (child.localName === "tbl") blocks.push(parseTable(child, context, comments));
      }
      applyContextualSpacing(blocks);
      const marginValue = (side: string): number | undefined => {
        const item = first(cellMargins ?? cell, side);
        return item ? twips(attr(item, "w"), 0) : tableMargins[side as keyof typeof tableMargins];
      };
      return {
        blocks,
        gridSpan: Math.max(1, span),
        rowSpan: val === "restart" ? -1 : val === "continue" || merge ? 0 : 1,
        width: cssWidth(tcW),
        shading: (shading && shading !== "auto" ? color(shading) : undefined) ?? style.cellShading,
        borders,
        margins: {
          top: marginValue("top"),
          right: marginValue("right"),
          bottom: marginValue("bottom"),
          left: marginValue("left"),
        },
        verticalAlign: attr(first(tcPr ?? cell, "vAlign"), "val"),
      };
    });
    return {
      cells,
      header: !!first(trPr ?? row, "tblHeader"),
      height: twips(attr(first(trPr ?? row, "trHeight"), "val"), 0),
      change: first(trPr ?? row, "ins")
        ? changeFrom(first(trPr ?? row, "ins")!, "ins")
        : first(trPr ?? row, "del")
          ? changeFrom(first(trPr ?? row, "del")!, "del")
          : undefined,
    };
  });

  const gridStarts = rows.map((row) => {
    let column = 0;
    return row.cells.map((cell) => {
      const start = column;
      column += cell.gridSpan;
      return start;
    });
  });
  for (let row = 0; row < rows.length; row++) {
    for (let cell = 0; cell < rows[row].cells.length; cell++) {
      const item = rows[row].cells[cell];
      if (item.rowSpan !== -1) continue;
      let span = 1;
      const start = gridStarts[row][cell];
      for (let next = row + 1; next < rows.length; next++) {
        const index = gridStarts[next].findIndex((column, candidate) => {
          const continuation = rows[next].cells[candidate];
          return (
            column === start &&
            continuation.rowSpan === 0 &&
            continuation.gridSpan === item.gridSpan
          );
        });
        if (index < 0) break;
        const continuation = rows[next].cells[index];
        continuation.rowSpan = -2;
        span++;
      }
      item.rowSpan = span;
    }
  }
  for (const row of rows) for (const cell of row.cells) if (cell.rowSpan === -2) cell.rowSpan = 0;

  const tableBorders: NonNullable<Table["borders"]> = { ...style.borders };
  const tblBorders = first(tblPr ?? node, "tblBorders");
  for (const side of ["top", "right", "bottom", "left", "insideH", "insideV"] as const) {
    const value = border(first(tblBorders ?? node, side));
    if (value !== undefined) tableBorders[side] = value;
  }
  return {
    type: "table",
    rows,
    widths,
    width: cssWidth(first(tblPr ?? node, "tblW")),
    indent: twips(attr(first(tblPr ?? node, "tblInd"), "w"), 0),
    align: attr(first(tblPr ?? node, "jc"), "val"),
    borders: tableBorders,
    margins: tableMargins,
  };
}
