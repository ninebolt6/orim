import type { Cell, CellRange, CellStyle, Worksheet } from "./model";

export const MAX_ROWS = 5000;
export const MAX_COLUMNS = 200;
export const HEADER_HEIGHT = 24;
export const HEADER_WIDTH = 48;

export interface LayoutRow {
  readonly index: number;
  readonly height: number;
  readonly stickyTop?: number;
}

export interface LayoutColumn {
  readonly index: number;
  readonly width: number;
  readonly style?: CellStyle;
  readonly stickyLeft?: number;
}

export interface LayoutMerge {
  readonly range: CellRange;
  readonly anchorRow: number;
  readonly anchorColumn: number;
  readonly rowSpan: number;
  readonly columnSpan: number;
  readonly cell?: Cell;
}

export interface WorksheetLayout {
  readonly rows: readonly LayoutRow[];
  readonly columns: readonly LayoutColumn[];
  readonly merges: readonly LayoutMerge[];
  readonly truncated: boolean;
  readonly rowAt: ReadonlyMap<number, LayoutRow>;
  readonly columnAt: ReadonlyMap<number, LayoutColumn>;
  readonly mergeAt: ReadonlyMap<string, LayoutMerge>;
  readonly covered: ReadonlySet<string>;
}

const key = (row: number, column: number): string => `${row}:${column}`;

const extendsExtent = (cell: Cell): boolean => {
  if (cell.value.type !== "empty") {
    if (cell.value.type !== "string" || cell.value.value.length > 0) return true;
  }
  return Boolean(cell.style.fill || Object.keys(cell.style.border).length);
};

const cellMap = (worksheet: Worksheet): Map<string, Cell> => {
  const result = new Map<string, Cell>();
  for (const row of worksheet.rows)
    for (const cell of row.cells) result.set(key(row.index, cell.column), cell);
  return result;
};

const columnRange = (worksheet: Worksheet, column: number) => {
  return worksheet.columns.find((range) => range.min <= column && column <= range.max);
};

export const computeLayout = (worksheet: Worksheet): WorksheetLayout => {
  const rowsByIndex = new Map(worksheet.rows.map((row) => [row.index, row]));
  const cells = cellMap(worksheet);
  let bottom = Math.max(-1, worksheet.frozen.rows - 1);
  let right = Math.max(-1, worksheet.frozen.columns - 1);
  for (const range of worksheet.merges) {
    bottom = Math.max(bottom, range.bottom);
    right = Math.max(right, range.right);
  }
  for (const row of worksheet.rows) {
    for (const cell of row.cells) {
      if (extendsExtent(cell)) {
        bottom = Math.max(bottom, row.index);
        right = Math.max(right, cell.column);
      }
    }
  }
  let truncated = false;
  const rows: LayoutRow[] = [];
  for (let index = 0; index <= bottom; index++) {
    if (rowsByIndex.get(index)?.hidden) continue;
    if (rows.length >= MAX_ROWS) {
      truncated = true;
      break;
    }
    const source = rowsByIndex.get(index);
    rows.push({
      index,
      height: source?.height ?? worksheet.defaultRowHeight,
    });
  }
  const columns: LayoutColumn[] = [];
  for (let index = 0; index <= right; index++) {
    const source = columnRange(worksheet, index);
    if (source?.hidden) continue;
    if (columns.length >= MAX_COLUMNS) {
      truncated = true;
      break;
    }
    columns.push({
      index,
      width: source?.width ?? worksheet.defaultColumnWidth,
      ...(source?.style ? { style: source.style } : {}),
    });
  }
  const visibleRows = new Map(rows.map((row, index) => [row.index, index]));
  const visibleColumns = new Map(columns.map((column, index) => [column.index, index]));
  const rowAt = new Map(rows.map((row) => [row.index, row]));
  const columnAt = new Map(columns.map((column) => [column.index, column]));
  let topOffset = worksheet.showHeaders ? HEADER_HEIGHT : 0;
  for (const row of rows) {
    if (row.index < worksheet.frozen.rows) {
      const frozenRow = { ...row, stickyTop: topOffset };
      rowAt.set(row.index, frozenRow);
      rows[visibleRows.get(row.index)!] = frozenRow;
      topOffset += (row.height * 4) / 3;
    }
  }
  let leftOffset = worksheet.showHeaders ? HEADER_WIDTH : 0;
  for (const column of columns) {
    if (column.index < worksheet.frozen.columns) {
      const frozenColumn = { ...column, stickyLeft: leftOffset };
      columnAt.set(column.index, frozenColumn);
      columns[visibleColumns.get(column.index)!] = frozenColumn;
      leftOffset += column.width;
    }
  }
  const merges: LayoutMerge[] = [];
  const mergeAt = new Map<string, LayoutMerge>();
  const covered = new Set<string>();
  const occupied = new Set<string>();
  for (const range of worksheet.merges) {
    const rowIndexes = rows
      .filter((row) => row.index >= range.top && row.index <= range.bottom)
      .map((row) => row.index);
    const columnIndexes = columns
      .filter((column) => column.index >= range.left && column.index <= range.right)
      .map((column) => column.index);
    const anchorRow = rowIndexes[0];
    const anchorColumn = columnIndexes[0];
    if (anchorRow === undefined || anchorColumn === undefined) continue;
    let overlaps = false;
    for (const row of rowIndexes)
      for (const column of columnIndexes) if (occupied.has(key(row, column))) overlaps = true;
    if (overlaps) continue;
    const merge: LayoutMerge = {
      range,
      anchorRow,
      anchorColumn,
      rowSpan: rowIndexes.length,
      columnSpan: columnIndexes.length,
      cell: cells.get(key(range.top, range.left)),
    };
    merges.push(merge);
    mergeAt.set(key(anchorRow, anchorColumn), merge);
    for (const row of rowIndexes)
      for (const column of columnIndexes)
        if (row !== anchorRow || column !== anchorColumn) covered.add(key(row, column));
    for (const row of rowIndexes)
      for (const column of columnIndexes) occupied.add(key(row, column));
  }
  return { rows, columns, merges, truncated, rowAt, columnAt, mergeAt, covered };
};

export const layoutKey = key;
