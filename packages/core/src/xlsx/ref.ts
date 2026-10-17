import type { CellRange } from "./model";

export const columnName = (index: number): string => {
  if (!Number.isInteger(index) || index < 0) return "";
  let value = index + 1;
  let name = "";
  while (value > 0) {
    value--;
    name = String.fromCharCode(65 + (value % 26)) + name;
    value = Math.floor(value / 26);
  }
  return name;
};

export const parseCellRef = (reference: string): { row: number; column: number } | undefined => {
  const match = /^\$?([A-Za-z]{1,3})\$?([1-9]\d*)$/.exec(reference);
  if (!match) return undefined;
  let column = 0;
  for (const character of match[1]!.toUpperCase())
    column = column * 26 + character.charCodeAt(0) - 64;
  column--;
  const row = Number(match[2]) - 1;
  if (!Number.isSafeInteger(row) || row >= 1048576 || column < 0 || column >= 16384)
    return undefined;
  return { row, column };
};

export const parseRange = (reference: string): CellRange | undefined => {
  const parts = reference.split(":");
  if (parts.length < 1 || parts.length > 2) return undefined;
  const first = parseCellRef(parts[0]!);
  const last = parseCellRef(parts[1] ?? parts[0]!);
  if (!first || !last) return undefined;
  return {
    top: Math.min(first.row, last.row),
    left: Math.min(first.column, last.column),
    bottom: Math.max(first.row, last.row),
    right: Math.max(first.column, last.column),
  };
};
