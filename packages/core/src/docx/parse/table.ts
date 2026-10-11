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
  const styleId = attr(first(tblPr, "tblStyle"), "val");
  const style = context.styles.table(styleId);
  const grid = first(node, "tblGrid");
  const widths = elements(grid, "gridCol").map((col) => cssWidth(col));
  const directTableMargins = first(tblPr, "tblCellMar");
  const sides = ["top", "right", "bottom", "left"] as const;
  const directMargins = Object.fromEntries(
    sides.flatMap((side) => {
      const item = first(directTableMargins, side);
      const value = item ? twips(attr(item, "w"), 0) : undefined;
      return value === undefined ? [] : [[side, value] as const];
    }),
  );
  const tableMargins = { ...style.margins, ...directMargins };
  const rows: TableRow[] = elements(node, "tr").map((row) => {
    const trPr = first(row, "trPr");
    const cells: TableCell[] = elements(row, "tc").map((cell) => {
      const tcPr = first(cell, "tcPr");
      const span = Number(attr(first(tcPr, "gridSpan"), "val") ?? 1);
      const merge = first(tcPr, "vMerge");
      const val = attr(merge, "val");
      const shading = attr(first(tcPr, "shd"), "fill");
      const tcW = first(tcPr, "tcW");
      const cellMargins = first(tcPr, "tcMar");
      const cellBorders = first(tcPr, "tcBorders");
      const borders = Object.fromEntries(
        sides.flatMap((side) => {
          const value = border(first(cellBorders, side));
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
        const item = first(cellMargins, side);
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
        verticalAlign: attr(first(tcPr, "vAlign"), "val"),
      };
    });
    return {
      cells,
      header: !!first(trPr, "tblHeader"),
      height: twips(attr(first(trPr, "trHeight"), "val"), 0),
      change: first(trPr, "ins")
        ? changeFrom(first(trPr, "ins")!, "ins")
        : first(trPr, "del")
          ? changeFrom(first(trPr, "del")!, "del")
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
  const tblBorders = first(tblPr, "tblBorders");
  const borderSides = [...sides, "insideH", "insideV"] as const;
  const tableBorders = Object.fromEntries(
    borderSides.flatMap((side) => {
      const value = border(first(tblBorders, side)) ?? style.borders[side];
      return value === undefined ? [] : [[side, value] as const];
    }),
  );
  return {
    type: "table",
    rows: mergedRows,
    widths,
    width: cssWidth(first(tblPr, "tblW")),
    indent: twips(attr(first(tblPr, "tblInd"), "w"), 0),
    align: attr(first(tblPr, "jc"), "val"),
    borders: tableBorders,
    margins: tableMargins,
  };
}
