import { hasValue, type Cell, type CellRange, type CellStyle, type Worksheet } from "./model";

export const MAX_ROWS = 5000;
export const MAX_COLUMNS = 200;
export const MAX_CELLS = 20_000;
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

export const visibleSpan = (
  items: readonly { readonly index: number }[],
  start: number,
  end: number,
): readonly [from: number, to: number] => {
  let low = 0;
  let high = items.length;
  while (low < high) {
    const middle = (low + high) >>> 1;
    if (items[middle]!.index < start) low = middle + 1;
    else high = middle;
  }
  const from = low;
  high = items.length;
  while (low < high) {
    const middle = (low + high) >>> 1;
    if (items[middle]!.index <= end) low = middle + 1;
    else high = middle;
  }
  return [from, low];
};

const key = (row: number, column: number): string => `${row}:${column}`;

const hasStyledExtent = (style: CellStyle | undefined): boolean =>
  Boolean(style?.fill || (style && Object.keys(style.border).length));

const extendsExtent = (cell: Cell): boolean => hasValue(cell.value) || hasStyledExtent(cell.style);

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
  let styledExtent = false;
  for (const range of worksheet.merges) {
    bottom = Math.max(bottom, range.bottom);
    right = Math.max(right, range.right);
  }
  for (const row of worksheet.rows) {
    if (hasStyledExtent(row.style)) {
      styledExtent = true;
      bottom = Math.max(bottom, row.index);
    }
    for (const cell of row.cells) {
      if (extendsExtent(cell)) {
        bottom = Math.max(bottom, row.index);
        right = Math.max(right, cell.column);
      }
    }
  }
  for (const range of worksheet.columns) {
    if (range.max < 16383 && hasStyledExtent(range.style)) {
      styledExtent = true;
      right = Math.max(right, range.max);
    }
  }
  if (styledExtent && bottom >= 0 && right < 0) right = 0;
  if (styledExtent && right >= 0 && bottom < 0) bottom = 0;
  let truncated = false;
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
  const maxRows = Math.min(
    MAX_ROWS,
    Math.max(1, Math.floor(MAX_CELLS / Math.max(1, columns.length))),
  );
  const rows: LayoutRow[] = [];
  for (let index = 0; index <= bottom; index++) {
    if (rowsByIndex.get(index)?.hidden) continue;
    if (rows.length >= maxRows) {
      truncated = true;
      break;
    }
    const source = rowsByIndex.get(index);
    rows.push({
      index,
      height: source?.height ?? worksheet.defaultRowHeight,
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
      .slice(...visibleSpan(rows, range.top, range.bottom))
      .map((row) => row.index);
    const columnIndexes = columns
      .slice(...visibleSpan(columns, range.left, range.right))
      .map((column) => column.index);
    const anchorRow = rowIndexes[0];
    const anchorColumn = columnIndexes[0];
    if (anchorRow === undefined || anchorColumn === undefined) continue;
    let overlaps = false;
    for (const row of rowIndexes) {
      for (const column of columnIndexes) {
        if (occupied.has(key(row, column))) {
          overlaps = true;
          break;
        }
      }
      if (overlaps) break;
    }
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
