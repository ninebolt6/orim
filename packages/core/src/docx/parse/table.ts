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

type MergeCell = Pick<TableCell, "gridSpan" | "rowSpan">;
type MergeRow = { readonly cells: readonly MergeCell[] };

export function computeRowSpans(rows: readonly MergeRow[]): number[][] {
  const gridStarts = rows.map((row) =>
    row.cells.map((_, index) =>
      row.cells.slice(0, index).reduce((column, cell) => column + cell.gridSpan, 0),
    ),
  );
  return rows.map((row, rowIndex) =>
    row.cells.map((cell: MergeCell, cellIndex) => {
      if (cell.rowSpan !== -1) return cell.rowSpan;
      const start = gridStarts[rowIndex][cellIndex];
      const continues = rows
        .slice(rowIndex + 1)
        .map((nextRow, offset) =>
          nextRow.cells.some(
            (candidate, candidateIndex) =>
              gridStarts[rowIndex + offset + 1][candidateIndex] === start &&
              candidate.rowSpan === 0 &&
              candidate.gridSpan === cell.gridSpan,
          ),
        );
      const interrupted = continues.findIndex((continuesMerge) => !continuesMerge);
      return 1 + (interrupted < 0 ? continues.length : interrupted);
    }),
  );
}

export function parseTable(node: Element, context: ParseContext): Table {
  const tblPr = first(node, "tblPr");
  const styleId = attr(first(tblPr ?? node, "tblStyle"), "val");
  const style = context.styles.table(styleId);
  const grid = first(node, "tblGrid");
  const widths = elements(grid ?? node, "gridCol").map((col) => cssWidth(col));
  const directTableMargins = first(tblPr ?? node, "tblCellMar");
  const sides = ["top", "right", "bottom", "left"] as const;
  const directMargins = Object.fromEntries(
    sides.flatMap((side) => {
      const item = first(directTableMargins ?? node, side);
      const value = item ? twips(attr(item, "w"), 0) : undefined;
      return value === undefined ? [] : [[side, value] as const];
    }),
  );
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
      const cellBorders = first(tcPr ?? cell, "tcBorders");
      const borders = Object.fromEntries(
        sides.flatMap((side) => {
          const value = border(first(cellBorders ?? cell, side));
          return value === undefined ? [] : [[side, value] as const];
        }),
      );
      const blocks = applyContextualSpacing(
        blockNodes(children(cell)).flatMap((child): Block[] =>
          child.localName === "p"
            ? splitAtPageBreaks(parseParagraph(child, context))
            : child.localName === "tbl"
              ? [parseTable(child, context)]
              : [],
        ),
      );
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

  const rowSpans = computeRowSpans(rows);
  const mergedRows = rows.map((row, rowIndex) => ({
    ...row,
    cells: row.cells.map((cell, cellIndex) => ({
      ...cell,
      rowSpan: rowSpans[rowIndex][cellIndex],
    })),
  }));
  const tblBorders = first(tblPr ?? node, "tblBorders");
  const borderSides = [...sides, "insideH", "insideV"] as const;
  const tableBorders = Object.fromEntries(
    borderSides.flatMap((side) => {
      const value = border(first(tblBorders ?? node, side)) ?? style.borders[side];
      return value === undefined ? [] : [[side, value] as const];
    }),
  );
  return {
    type: "table",
    rows: mergedRows,
    widths,
    width: cssWidth(first(tblPr ?? node, "tblW")),
    indent: twips(attr(first(tblPr ?? node, "tblInd"), "w"), 0),
    align: attr(first(tblPr ?? node, "jc"), "val"),
    borders: tableBorders,
    margins: tableMargins,
  };
}
