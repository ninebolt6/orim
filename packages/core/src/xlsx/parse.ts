import { attr, children, elements, first, R } from "../common/xml";
import { safeHref } from "../common/sanitize";
import { openOpc, type OpcPackage, type Relationship } from "../opc/package";
import type {
  Cell,
  CellValue,
  Hyperlink,
  SheetRow,
  SheetState,
  TextRun,
  Worksheet,
  XlsxWorkbook,
} from "./model";
import { parseCellRef, parseRange } from "./ref";
import { themeColors } from "./colors";
import { defaultCellStyle, parseStyles } from "./parse/styles";
import type { StyleTable } from "./parse/styles";

const EMPTY: CellValue = { type: "empty" };
const maxRows = 1048576;
const maxColumns = 16384;

const relationship = (
  relationships: ReadonlyMap<string, Relationship>,
  suffix: string,
): Relationship | undefined => {
  return Array.from(relationships.values()).find(
    (item) => !item.external && item.type.endsWith(`/${suffix}`),
  );
};

const relationshipId = (element: Element | undefined): string | undefined => {
  if (!element) return undefined;
  const direct = attr(element, "id", R);
  if (direct) return direct;
  return Array.from(element.attributes).find(
    (item) => item.localName === "id" && item.prefix === "r",
  )?.value;
};

const boolAttribute = (element: Element | undefined, name: string, fallback: boolean): boolean => {
  const value = attr(element, name);
  if (value === undefined) return fallback;
  return !["0", "false", "off", "no"].includes(value.toLowerCase());
};

const integerAttribute = (element: Element | undefined, name: string): number | undefined => {
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
      value += decodeXstring(child.textContent ?? "");
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
  shared: readonly { readonly value: string; readonly runs?: readonly TextRun[] }[],
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
  return Math.trunc(((256 * width + Math.trunc(128 / 7)) / 256) * 7);
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

const parseWorksheet = async (
  pkg: OpcPackage,
  path: string,
  name: string,
  state: SheetState,
  date1904: boolean,
  shared: readonly { readonly value: string; readonly runs?: readonly TextRun[] }[],
  styles: StyleTable,
): Promise<Worksheet> => {
  const document = await pkg.xml(path);
  const root = document.documentElement;
  const sheetFormat = first(root, "sheetFormatPr");
  const rawDefaultColumn = Number(attr(sheetFormat, "defaultColWidth"));
  const baseColumn = Number(attr(sheetFormat, "baseColWidth"));
  const defaultColumnWidth =
    Number.isFinite(rawDefaultColumn) && rawDefaultColumn >= 0
      ? widthInPixels(rawDefaultColumn)
      : (Number.isFinite(baseColumn) && baseColumn >= 0 ? baseColumn : 8) * 7 + 5;
  const rawDefaultRow = Number(attr(sheetFormat, "defaultRowHeight"));
  const defaultRowHeight =
    Number.isFinite(rawDefaultRow) && rawDefaultRow >= 0 ? rawDefaultRow : 15;
  const columns = elements(first(root, "cols"), "col").flatMap((column) => {
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
    if (Number.isFinite(height) && height >= 0) row.height = height;
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
  const sheetRelationships = await pkg.relationships(path);
  const hyperlinks: Hyperlink[] = [];
  for (const link of elements(first(root, "hyperlinks"), "hyperlink")) {
    const range = parseRange(attr(link, "ref") ?? "");
    const relId = relationshipId(link);
    const rel = relId ? sheetRelationships.get(relId) : undefined;
    const href = rel?.external ? safeHref(rel.target) : undefined;
    if (range && href && rel?.type.endsWith("/hyperlink")) hyperlinks.push({ range, href });
  }
  const sheetView = elements(first(root, "sheetViews"), "sheetView")[0];
  const pane = first(sheetView, "pane");
  const frozenState = attr(pane, "state");
  const frozen = frozenState === "frozen" || frozenState === "frozenSplit";
  const frozenRows =
    frozen && Number.isFinite(Number(attr(pane, "ySplit")))
      ? Math.min(maxRows, Math.max(0, Math.trunc(Number(attr(pane, "ySplit")))))
      : 0;
  const frozenColumns =
    frozen && Number.isFinite(Number(attr(pane, "xSplit")))
      ? Math.min(maxColumns, Math.max(0, Math.trunc(Number(attr(pane, "xSplit")))))
      : 0;
  const merges = elements(first(root, "mergeCells"), "mergeCell").flatMap((merge) => {
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

const parseSharedStrings = async (
  pkg: OpcPackage,
  path: string | undefined,
  styles: StyleTable,
): Promise<readonly { readonly value: string; readonly runs?: readonly TextRun[] }[]> => {
  if (!path) return [];
  const document = await pkg.xml(path);
  return elements(document.documentElement, "si").map((item) => richText(item, styles));
};

export const parseXlsx = async (input: ArrayBuffer | Uint8Array | Blob): Promise<XlsxWorkbook> => {
  let bytes: Uint8Array;
  if (typeof Blob !== "undefined" && input instanceof Blob)
    bytes = new Uint8Array(await input.arrayBuffer());
  else if (input instanceof Uint8Array) bytes = input;
  else bytes = new Uint8Array(input as ArrayBuffer);
  const pkg = openOpc(bytes);
  const rootRelationships = await pkg.relationships("");
  const workbookRelationship = relationship(rootRelationships, "officeDocument");
  if (!workbookRelationship) throw new Error("Not an XLSX package: workbook part not found.");
  const workbookPath = pkg.resolve("", workbookRelationship.target);
  const workbookDocument = await pkg.xml(workbookPath);
  const workbookRelationships = await pkg.relationships(workbookPath);
  const stylesPath = relationship(workbookRelationships, "styles");
  const sharedPath = relationship(workbookRelationships, "sharedStrings");
  const themePath = relationship(workbookRelationships, "theme");
  const themeDocument = themePath
    ? await pkg.xml(pkg.resolve(workbookPath, themePath.target))
    : undefined;
  const theme = themeColors(themeDocument);
  const styleDocument = stylesPath
    ? await pkg.xml(pkg.resolve(workbookPath, stylesPath.target))
    : undefined;
  const styles = parseStyles(styleDocument, theme);
  const shared = await parseSharedStrings(
    pkg,
    sharedPath ? pkg.resolve(workbookPath, sharedPath.target) : undefined,
    styles,
  );
  const workbookRoot = workbookDocument.documentElement;
  const workbookProperties = first(workbookRoot, "workbookPr");
  const date1904 = boolAttribute(workbookProperties, "date1904", false);
  const sheetEntries = elements(first(workbookRoot, "sheets"), "sheet");
  const parsedSheets = await Promise.all(
    sheetEntries.map(async (sheet) => {
      const id = relationshipId(sheet);
      const rel = id ? workbookRelationships.get(id) : undefined;
      if (!rel || rel.external || !rel.type.endsWith("/worksheet")) return undefined;
      const state = attr(sheet, "state");
      return parseWorksheet(
        pkg,
        pkg.resolve(workbookPath, rel.target),
        attr(sheet, "name") ?? "",
        state === "hidden" || state === "veryHidden" ? state : "visible",
        date1904,
        shared,
        styles,
      );
    }),
  );
  const sheets: Worksheet[] = [];
  const sheetIndexes = new Map<number, number>();
  parsedSheets.forEach((sheet, index) => {
    if (sheet) {
      sheetIndexes.set(index, sheets.length);
      sheets.push(sheet);
    }
  });
  const workbookView = elements(first(workbookRoot, "bookViews"), "workbookView")[0];
  const activeTab = integerAttribute(workbookView, "activeTab") ?? 0;
  const mappedActive = sheetIndexes.get(activeTab);
  const firstVisible = sheets.findIndex((sheet) => sheet.state === "visible");
  const activeSheet = mappedActive ?? (firstVisible >= 0 ? firstVisible : 0);
  return { sheets, activeSheet, date1904 };
};

export { defaultCellStyle };
