import { attr, children, elements, first, R } from "../../common/xml";
import { safeHref } from "../../common/sanitize";
import type { Relationship } from "../../opc/package";
import type {
  Cell,
  CellValue,
  Hyperlink,
  SheetRow,
  SheetState,
  TextRun,
  Worksheet,
} from "../model";
import { parseCellRef, parseRange } from "../ref";
import { MAX_CELLS, MAX_COLUMNS, MAX_ROWS } from "../layout";
import type { StyleTable } from "./styles";

const EMPTY: CellValue = { type: "empty" };
const maxRows = 1048576;
const maxColumns = 16384;
const maxColumnWidth = 255;
const maxRowHeight = 409;

export interface SharedString {
  readonly value: string;
  readonly runs?: readonly TextRun[];
}

export interface WorksheetContext {
  readonly name: string;
  readonly state: SheetState;
  readonly date1904: boolean;
  readonly shared: readonly SharedString[];
  readonly styles: StyleTable;
  readonly relationships: ReadonlyMap<string, Relationship>;
}

export const relationshipId = (element: Element | undefined): string | undefined => {
  if (!element) return undefined;
  const direct = attr(element, "id", R);
  if (direct) return direct;
  return Array.from(element.attributes).find(
    (item) => item.localName === "id" && item.prefix === "r",
  )?.value;
};

export const boolAttribute = (
  element: Element | undefined,
  name: string,
  fallback: boolean,
): boolean => {
  const value = attr(element, name);
  if (value === undefined) return fallback;
  return !["0", "false", "off", "no"].includes(value.toLowerCase());
};

export const integerAttribute = (
  element: Element | undefined,
  name: string,
): number | undefined => {
  const value = Number(attr(element, name));
  return Number.isInteger(value) ? value : undefined;
};

const decodeXstring = (value: string): string => {
  return value.replace(/_x([0-9a-f]{4})_/gi, (escape, digits: string) => {
    return String.fromCharCode(Number.parseInt(digits, 16));
  });
};

const richText = (
  element: Element | undefined,
  styles: StyleTable,
): { value: string; runs?: readonly TextRun[] } => {
  if (!element) return { value: "" };
  let value = "";
  let hasRuns = false;
  const runs: TextRun[] = [];
  for (const child of children(element)) {
    if (child.localName === "t") {
      const runText = decodeXstring(child.textContent ?? "");
      value += runText;
      runs.push({ text: runText });
    } else if (child.localName === "r") {
      hasRuns = true;
      const runText = decodeXstring(first(child, "t")?.textContent ?? "");
      value += runText;
      const font = styles.fontForRun(first(child, "rPr"));
      runs.push(font ? { text: runText, font } : { text: runText });
    }
  }
  return hasRuns ? { value, runs } : { value };
};

const dateToSerial = (value: string, date1904: boolean): number | undefined => {
  const iso = value.trim();
  const withUtc = /(?:Z|[+-]\d{2}:\d{2})$/i.test(iso) || !iso.includes("T") ? iso : `${iso}Z`;
  const timestamp = Date.parse(withUtc);
  if (!Number.isFinite(timestamp)) return undefined;
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 31);
  let serial = (timestamp - epoch) / 86400000;
  if (!date1904 && timestamp >= Date.UTC(1900, 2, 1)) serial++;
  return Number.isFinite(serial) ? serial : undefined;
};

const cellValue = (
  cell: Element,
  shared: readonly SharedString[],
  date1904: boolean,
  styles: StyleTable,
): { value: CellValue; runs?: readonly TextRun[] } => {
  const type = attr(cell, "t") ?? "n";
  if (type === "inlineStr") {
    const result = richText(first(cell, "is"), styles);
    return result.value
      ? { value: { type: "string", value: result.value }, runs: result.runs }
      : { value: EMPTY };
  }
  const raw = first(cell, "v")?.textContent ?? "";
  if (!raw.trim()) return { value: EMPTY };
  if (type === "s") {
    const index = Number(raw);
    const item = Number.isInteger(index) ? shared[index] : undefined;
    return item
      ? {
          value: { type: "string", value: item.value },
          ...(item.runs ? { runs: item.runs } : {}),
        }
      : { value: EMPTY };
  }
  if (type === "str") return { value: { type: "string", value: decodeXstring(raw) } };
  if (type === "b") {
    if (raw === "1") return { value: { type: "boolean", value: true } };
    if (raw === "0") return { value: { type: "boolean", value: false } };
    return { value: EMPTY };
  }
  if (type === "e") return { value: { type: "error", value: raw } };
  if (type === "d") {
    const serial = dateToSerial(raw, date1904);
    return serial === undefined ? { value: EMPTY } : { value: { type: "number", value: serial } };
  }
  if (type === "n") {
    const numeric = Number(raw);
    return Number.isFinite(numeric)
      ? { value: { type: "number", value: numeric } }
      : { value: EMPTY };
  }
  return { value: EMPTY };
};

const widthInPixels = (width: number): number => {
  const bounded = Math.min(maxColumnWidth, width);
  return Math.trunc(((256 * bounded + Math.trunc(128 / 7)) / 256) * 7);
};

const rowFor = (
  rows: Map<
    number,
    {
      height?: number;
      hidden: boolean;
      style?: ReturnType<StyleTable["resolve"]>;
      cells: Map<number, Cell>;
    }
  >,
  index: number,
): {
  height?: number;
  hidden: boolean;
  style?: ReturnType<StyleTable["resolve"]>;
  cells: Map<number, Cell>;
} => {
  let row = rows.get(index);
  if (!row) {
    row = { hidden: false, cells: new Map() };
    rows.set(index, row);
  }
  return row;
};

export const parseWorksheet = (root: Element, context: WorksheetContext): Worksheet => {
  const { name, state, date1904, shared, styles, relationships } = context;
  const sheetFormat = first(root, "sheetFormatPr");
  const rawDefaultColumn = Number(attr(sheetFormat, "defaultColWidth"));
  const baseColumn = Number(attr(sheetFormat, "baseColWidth"));
  const defaultColumnWidth =
    Number.isFinite(rawDefaultColumn) && rawDefaultColumn >= 0
      ? widthInPixels(rawDefaultColumn)
      : Math.min(maxColumnWidth, Number.isFinite(baseColumn) && baseColumn >= 0 ? baseColumn : 8) *
          7 +
        5;
  const rawDefaultRow = Number(attr(sheetFormat, "defaultRowHeight"));
  const defaultRowHeight =
    Number.isFinite(rawDefaultRow) && rawDefaultRow >= 0
      ? Math.min(maxRowHeight, rawDefaultRow)
      : 15;
  const columns = elements(first(root, "cols"), "col")
    .slice(0, MAX_CELLS)
    .flatMap((column) => {
      const min = Number(attr(column, "min")) - 1;
      const max = Number(attr(column, "max")) - 1;
      if (
        !Number.isInteger(min) ||
        !Number.isInteger(max) ||
        min < 0 ||
        max < min ||
        max >= maxColumns
      )
        return [];
      const rawWidth = Number(attr(column, "width"));
      const styleId = integerAttribute(column, "style");
      return [
        {
          min,
          max,
          ...(Number.isFinite(rawWidth) && rawWidth >= 0 ? { width: widthInPixels(rawWidth) } : {}),
          hidden: boolAttribute(column, "hidden", false),
          ...(styleId !== undefined ? { style: styles.resolve(styleId) } : {}),
        },
      ];
    });
  const rows = new Map<
    number,
    {
      height?: number;
      hidden: boolean;
      style?: ReturnType<StyleTable["resolve"]>;
      cells: Map<number, Cell>;
    }
  >();
  let previousRow = -1;
  for (const rowElement of elements(first(root, "sheetData"), "row")) {
    const rowReference = integerAttribute(rowElement, "r");
    const rowIndex = rowReference === undefined ? previousRow + 1 : rowReference - 1;
    previousRow = rowIndex;
    if (rowIndex < 0 || rowIndex >= maxRows) continue;
    const row = rowFor(rows, rowIndex);
    const height = Number(attr(rowElement, "ht"));
    if (Number.isFinite(height) && height >= 0) row.height = Math.min(maxRowHeight, height);
    row.hidden = boolAttribute(rowElement, "hidden", false);
    if (boolAttribute(rowElement, "customFormat", false)) {
      const styleId = integerAttribute(rowElement, "s");
      if (styleId !== undefined) row.style = styles.resolve(styleId);
    }
    let previousColumn = -1;
    for (const cellElement of elements(rowElement, "c")) {
      const reference = attr(cellElement, "r");
      const parsedReference = reference ? parseCellRef(reference) : undefined;
      if (reference && !parsedReference) continue;
      const column = parsedReference?.column ?? previousColumn + 1;
      const cellRow = parsedReference?.row ?? rowIndex;
      previousColumn = column;
      if (column < 0 || column >= maxColumns || cellRow < 0 || cellRow >= maxRows) continue;
      const targetRow = rowFor(rows, cellRow);
      const styleId = integerAttribute(cellElement, "s");
      const cellStyle = styles.resolve(styleId);
      const parsed = cellValue(cellElement, shared, date1904, styles);
      const cell: Cell = {
        column,
        value: parsed.value,
        style: cellStyle,
        ...(parsed.runs ? { runs: parsed.runs } : {}),
      };
      targetRow.cells.set(column, cell);
    }
  }
  const hyperlinks: Hyperlink[] = [];
  for (const link of elements(first(root, "hyperlinks"), "hyperlink").slice(0, MAX_CELLS)) {
    const range = parseRange(attr(link, "ref") ?? "");
    const relId = relationshipId(link);
    const rel = relId ? relationships.get(relId) : undefined;
    const href = rel?.external ? safeHref(rel.target) : undefined;
    if (range && href && rel?.type.endsWith("/hyperlink")) hyperlinks.push({ range, href });
  }
  const sheetView = elements(first(root, "sheetViews"), "sheetView")[0];
  const pane = first(sheetView, "pane");
  const frozenState = attr(pane, "state");
  const frozen = frozenState === "frozen" || frozenState === "frozenSplit";
  const frozenRows =
    frozen && Number.isFinite(Number(attr(pane, "ySplit")))
      ? Math.min(MAX_ROWS, Math.max(0, Math.trunc(Number(attr(pane, "ySplit")))))
      : 0;
  const frozenColumns =
    frozen && Number.isFinite(Number(attr(pane, "xSplit")))
      ? Math.min(MAX_COLUMNS, Math.max(0, Math.trunc(Number(attr(pane, "xSplit")))))
      : 0;
  const merges = elements(first(root, "mergeCells"), "mergeCell")
    .slice(0, MAX_CELLS)
    .flatMap((merge) => {
      const range = parseRange(attr(merge, "ref") ?? "");
      return range ? [range] : [];
    });
  return {
    name,
    state,
    rows: Array.from(rows, ([index, row]): SheetRow => ({
      index,
      ...(row.height !== undefined ? { height: row.height } : {}),
      hidden: row.hidden,
      ...(row.style ? { style: row.style } : {}),
      cells: Array.from(row.cells.values()).sort((left, right) => left.column - right.column),
    })).sort((left, right) => left.index - right.index),
    columns,
    merges,
    hyperlinks,
    defaultColumnWidth,
    defaultRowHeight,
    frozen: {
      rows: Number.isInteger(frozenRows) && frozenRows > 0 ? frozenRows : 0,
      columns: Number.isInteger(frozenColumns) && frozenColumns > 0 ? frozenColumns : 0,
    },
    showGridLines: boolAttribute(sheetView, "showGridLines", true),
    showHeaders: boolAttribute(sheetView, "showRowColHeaders", true),
  };
};

export const parseSharedStrings = (
  root: Element | undefined,
  styles: StyleTable,
): readonly SharedString[] => {
  if (!root) return [];
  return elements(root, "si").map((item) => richText(item, styles));
};
