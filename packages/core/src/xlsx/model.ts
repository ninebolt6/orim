export interface XlsxWorkbook {
  readonly sheets: readonly Worksheet[];
  readonly activeSheet: number;
  readonly date1904: boolean;
}

export type SheetState = "visible" | "hidden" | "veryHidden";

export interface Worksheet {
  readonly name: string;
  readonly state: SheetState;
  readonly rows: readonly SheetRow[];
  readonly columns: readonly ColumnRange[];
  readonly merges: readonly CellRange[];
  readonly hyperlinks: readonly Hyperlink[];
  readonly defaultColumnWidth: number;
  readonly defaultRowHeight: number;
  readonly frozen: { readonly rows: number; readonly columns: number };
  readonly showGridLines: boolean;
  readonly showHeaders: boolean;
}

export interface CellRange {
  readonly top: number;
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
}

export interface ColumnRange {
  readonly min: number;
  readonly max: number;
  readonly width?: number;
  readonly hidden: boolean;
  readonly style?: CellStyle;
}

export interface SheetRow {
  readonly index: number;
  readonly height?: number;
  readonly hidden: boolean;
  readonly style?: CellStyle;
  readonly cells: readonly Cell[];
}

export interface Cell {
  readonly column: number;
  readonly value: CellValue;
  readonly style: CellStyle;
  readonly runs?: readonly TextRun[];
}

export type CellValue =
  | { readonly type: "empty" }
  | { readonly type: "number"; readonly value: number }
  | { readonly type: "string"; readonly value: string }
  | { readonly type: "boolean"; readonly value: boolean }
  | { readonly type: "error"; readonly value: string };

export const hasValue = (value: CellValue): boolean =>
  value.type !== "empty" && (value.type !== "string" || value.value.length > 0);

export interface Hyperlink {
  readonly range: CellRange;
  readonly href: string;
}

export interface CellStyle {
  readonly numberFormat: string;
  readonly font: FontStyle;
  readonly fill?: string;
  readonly border: BorderSet;
  readonly alignment: Alignment;
}

export interface FontStyle {
  readonly name?: string;
  readonly size?: number;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly underline?: "single" | "double";
  readonly strike: boolean;
  readonly color?: string;
  readonly vertAlign?: "superscript" | "subscript";
}

export type BorderLine =
  | "thin"
  | "medium"
  | "thick"
  | "dashed"
  | "dotted"
  | "double"
  | "hair"
  | "mediumDashed"
  | "dashDot"
  | "mediumDashDot"
  | "dashDotDot"
  | "mediumDashDotDot"
  | "slantDashDot";

export interface BorderEdge {
  readonly line: BorderLine;
  readonly color?: string;
}

export interface BorderSet {
  readonly top?: BorderEdge;
  readonly right?: BorderEdge;
  readonly bottom?: BorderEdge;
  readonly left?: BorderEdge;
}

export interface Alignment {
  readonly horizontal?:
    | "left"
    | "center"
    | "right"
    | "fill"
    | "justify"
    | "centerContinuous"
    | "distributed";
  readonly vertical?: "top" | "center" | "bottom" | "justify" | "distributed";
  readonly wrap: boolean;
  readonly indent: number;
}

export interface TextRun {
  readonly text: string;
  readonly font?: Partial<FontStyle>;
}
