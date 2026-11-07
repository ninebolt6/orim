import { attr, children, descendants, elements, first, textValue } from "../../common/xml";
import { parseColor } from "../../drawingml/color";
import {
  fillElement,
  mergeLine,
  parseFill,
  parseLine,
  type Fill,
  type LineStyle,
} from "../../drawingml/fill";
import { emu } from "../../drawingml/units";
import type { CellBorders, RunStyle, TableCell, TableElement } from "../model";
import { defined, layer, type ShapeContext } from "./context";
import { fillOptions, referenceFill, referenceFont, referenceLine } from "./style";
import { parseTextBody, uniformListStyle, type BodyProps } from "./text";

type Edge = "left" | "right" | "top" | "bottom" | "insideH" | "insideV";

interface Region {
  readonly fill?: Fill;
  readonly borders: Partial<Record<Edge, LineStyle>>;
  readonly text: Partial<RunStyle>;
}

interface Range {
  readonly rows: readonly [number, number];
  readonly columns: readonly [number, number];
}

const edges: readonly Edge[] = ["left", "right", "top", "bottom", "insideH", "insideV"];

const on = (value: string | undefined): boolean | undefined =>
  value === undefined ? undefined : value === "on";

const parseRegion = (element: Element | undefined, context: ShapeContext): Region | undefined => {
  if (!element) return undefined;
  const cellStyle = first(element, "tcStyle");
  const fill = first(cellStyle, "fill");
  const borders = first(cellStyle, "tcBdr");
  const textStyle = first(element, "tcTxStyle");
  return {
    fill: fill
      ? parseFill(fillElement(fill), fillOptions(context))
      : referenceFill(first(cellStyle, "fillRef"), context),
    borders: Object.fromEntries(
      edges.flatMap((edge) => {
        const border = first(borders, edge);
        const line = first(border, "ln")
          ? parseLine(first(border, "ln"), fillOptions(context))
          : referenceLine(first(border, "lnRef"), context);
        return line ? [[edge, line] as const] : [];
      }),
    ),
    text: defined({
      ...referenceFont(first(textStyle, "fontRef"), context),
      bold: on(attr(textStyle, "b")),
      italic: on(attr(textStyle, "i")),
      ...(textStyle && defined({ color: parseColor(textStyle, context.scheme) })),
    }),
  };
};

const regionEdges = (
  region: Region,
  range: Range,
  row: readonly [number, number],
  column: readonly [number, number],
): CellBorders =>
  defined({
    left: column[0] === range.columns[0] ? region.borders.left : region.borders.insideV,
    right: column[1] === range.columns[1] ? region.borders.right : region.borders.insideV,
    top: row[0] === range.rows[0] ? region.borders.top : region.borders.insideH,
    bottom: row[1] === range.rows[1] ? region.borders.bottom : region.borders.insideH,
  });

const cellMargins = (properties: Element | undefined): BodyProps => {
  const anchor = attr(properties, "anchor");
  return defined({
    left: emu(attr(properties, "marL")) ?? 7.2,
    right: emu(attr(properties, "marR")) ?? 7.2,
    top: emu(attr(properties, "marT")) ?? 3.6,
    bottom: emu(attr(properties, "marB")) ?? 3.6,
    anchor: anchor === "ctr" ? "middle" : anchor === "b" ? "bottom" : "top",
  });
};

const explicitBorders = (properties: Element | undefined, context: ShapeContext): CellBorders =>
  defined({
    left: parseLine(first(properties, "lnL"), fillOptions(context)),
    right: parseLine(first(properties, "lnR"), fillOptions(context)),
    top: parseLine(first(properties, "lnT"), fillOptions(context)),
    bottom: parseLine(first(properties, "lnB"), fillOptions(context)),
  });

export const parseTable = (
  frame: Element,
  base: Pick<TableElement, "id" | "name" | "transform" | "href">,
  context: ShapeContext,
): TableElement | undefined => {
  const table = descendants(frame, "tbl")[0];
  if (!table) return undefined;
  const properties = first(table, "tblPr");
  const enabled = (name: string) => attr(properties, name) === "1";
  const style = context.tableStyles.get(textValue(first(properties, "tableStyleId")).trim());
  const region = (name: string) => parseRegion(first(style, name), context);
  const columns = elements(first(table, "tblGrid"), "gridCol").map(
    (column) => emu(attr(column, "w")) ?? 0,
  );
  const rowElements = elements(table, "tr");
  const lastRow = rowElements.length - 1;
  const lastColumn = columns.length - 1;
  const whole = region("wholeTbl");
  const bands = {
    band1H: region("band1H"),
    band2H: region("band2H"),
    band1V: region("band1V"),
    band2V: region("band2V"),
  };
  const firstRow = enabled("firstRow") ? region("firstRow") : undefined;
  const lastRowRegion = enabled("lastRow") ? region("lastRow") : undefined;
  const firstColumn = enabled("firstCol") ? region("firstCol") : undefined;
  const lastColumnRegion = enabled("lastCol") ? region("lastCol") : undefined;
  const allRows = [0, lastRow] as const;
  const allColumns = [0, lastColumn] as const;
  const regionsFor = (row: number, column: number): [Region | undefined, Range][] => {
    const bandRow = row - (firstRow ? 1 : 0);
    const bandColumn = column - (firstColumn ? 1 : 0);
    return [
      [whole, { rows: allRows, columns: allColumns }],
      [
        enabled("bandCol") ? (bandColumn % 2 ? bands.band2V : bands.band1V) : undefined,
        { rows: allRows, columns: [column, column] },
      ],
      [
        enabled("bandRow") ? (bandRow % 2 ? bands.band2H : bands.band1H) : undefined,
        { rows: [row, row], columns: allColumns },
      ],
      [
        column === lastColumn ? lastColumnRegion : undefined,
        { rows: allRows, columns: [column, column] },
      ],
      [column === 0 ? firstColumn : undefined, { rows: allRows, columns: [0, 0] }],
      [row === lastRow ? lastRowRegion : undefined, { rows: [row, row], columns: allColumns }],
      [row === 0 ? firstRow : undefined, { rows: [0, 0], columns: allColumns }],
    ];
  };
  const rows = rowElements.map((row, rowIndex) => {
    let columnIndex = 0;
    const cells = elements(row, "tc").map((cell): TableCell => {
      const column = columnIndex;
      const columnSpan = Math.max(1, Number(attr(cell, "gridSpan")) || 1);
      const rowSpan = Math.max(1, Number(attr(cell, "rowSpan")) || 1);
      columnIndex += 1;
      const span = [
        [rowIndex, Math.min(lastRow, rowIndex + rowSpan - 1)] as const,
        [column, Math.min(lastColumn, column + columnSpan - 1)] as const,
      ] as const;
      const regions = regionsFor(rowIndex, column).filter(
        (entry): entry is [Region, Range] => !!entry[0],
      );
      const cellProperties = first(cell, "tcPr");
      const borders = regions.reduce<CellBorders>((result, [item, range]) => {
        const next = regionEdges(item, range, span[0], span[1]);
        return {
          left: mergeLine(result.left, next.left),
          right: mergeLine(result.right, next.right),
          top: mergeLine(result.top, next.top),
          bottom: mergeLine(result.bottom, next.bottom),
        };
      }, {});
      const explicit = explicitBorders(cellProperties, context);
      const ownFill = fillElement(cellProperties);
      const text = layer<RunStyle>(...regions.map(([item]) => item.text));
      return defined({
        text: parseTextBody(
          first(cell, "txBody"),
          [context.defaultText, uniformListStyle(text)],
          [cellMargins(cellProperties)],
          context,
        ),
        fill: ownFill
          ? parseFill(ownFill, fillOptions(context))
          : regions.reduce<Fill | undefined>((result, [item]) => item.fill ?? result, undefined),
        borders: defined({
          left: mergeLine(borders.left, explicit.left),
          right: mergeLine(borders.right, explicit.right),
          top: mergeLine(borders.top, explicit.top),
          bottom: mergeLine(borders.bottom, explicit.bottom),
        }),
        columnSpan,
        rowSpan,
        merged: attr(cell, "hMerge") === "1" || attr(cell, "vMerge") === "1",
      }) as TableCell;
    });
    return { height: emu(attr(row, "h")) ?? 0, cells };
  });
  return { kind: "table", ...base, columns, rows };
};

const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const solid = (color: string) => `<a:solidFill><a:schemeClr val="${color}"/></a:solidFill>`;
const tinted = (tint: number) =>
  `<a:fill><a:solidFill><a:schemeClr val="accent1"><a:tint val="${tint}"/></a:schemeClr></a:solidFill></a:fill>`;
const border = (edge: string, width: number) =>
  `<a:${edge}><a:ln w="${width}">${solid("lt1")}</a:ln></a:${edge}>`;
const emphasis = (edge: string) =>
  `<a:tcTxStyle b="on"><a:schemeClr val="lt1"/></a:tcTxStyle><a:tcStyle>${
    edge ? `<a:tcBdr>${border(edge, 38100)}</a:tcBdr>` : ""
  }<a:fill>${solid("accent1")}</a:fill></a:tcStyle>`;

/** PowerPoint's default "Medium Style 2 – Accent 1", referenced but rarely embedded. */
export const builtinTableStyles = (): ReadonlyMap<string, Element> => {
  const id = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";
  const xml = `<a:tblStyle xmlns:a="${A}" styleId="${id}"><a:wholeTbl><a:tcTxStyle><a:schemeClr val="dk1"/></a:tcTxStyle><a:tcStyle><a:tcBdr>${[
    "left",
    "right",
    "top",
    "bottom",
    "insideH",
    "insideV",
  ]
    .map((edge) => border(edge, 12700))
    .join(
      "",
    )}</a:tcBdr>${tinted(20000)}</a:tcStyle></a:wholeTbl><a:band1H><a:tcStyle>${tinted(40000)}</a:tcStyle></a:band1H><a:band1V><a:tcStyle>${tinted(40000)}</a:tcStyle></a:band1V><a:lastCol>${emphasis("")}</a:lastCol><a:firstCol>${emphasis("")}</a:firstCol><a:lastRow>${emphasis("top")}</a:lastRow><a:firstRow>${emphasis("bottom")}</a:firstRow></a:tblStyle>`;
  const root = new DOMParser().parseFromString(xml, "application/xml").documentElement;
  return new Map([[id, root]]);
};

export const parseTableStyles = (doc: Document | undefined): ReadonlyMap<string, Element> =>
  new Map([
    ...builtinTableStyles(),
    ...(doc ? children(doc.documentElement) : []).map(
      (style) => [attr(style, "styleId") ?? "", style] as const,
    ),
  ]);
