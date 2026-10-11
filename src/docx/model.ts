export type ChangeKind = "ins" | "del" | "moveFrom" | "moveTo";

export interface ChangeMark {
  kind: ChangeKind;
  author: string;
  date: string;
  color?: string;
}

export interface TextStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: string;
  strike?: boolean;
  doubleStrike?: boolean;
  color?: string;
  highlight?: string;
  fontFamily?: string;
  fontSize?: number;
  verticalAlign?: "superscript" | "subscript";
  caps?: boolean;
  smallCaps?: boolean;
  hidden?: boolean;
  backgroundColor?: string;
  style?: Record<string, string>;
}

export type Inline =
  | {
      type: "text";
      text: string;
      style?: TextStyle;
      change?: ChangeMark;
      formattingChange?: ChangeMark;
      comments?: string[];
      href?: string;
    }
  | { type: "image"; path: string; width?: number; height?: number }
  | { type: "field"; name: string; cached?: string }
  | { type: "break"; page?: boolean }
  | { type: "bookmark"; name: string }
  | { type: "comment-reference"; id: string };

export interface ParagraphProperties {
  align?: string;
  styleId?: string;
  before?: number;
  after?: number;
  line?: number;
  lineRule?: string;
  left?: number;
  right?: number;
  firstLine?: number;
  hanging?: number;
  contextualSpacing?: boolean;
  pageBreakBefore?: boolean;
  keepNext?: boolean;
  shading?: string;
  border?: string;
  change?: ChangeMark;
  formattingChange?: ChangeMark;
}

export interface Paragraph {
  type: "paragraph";
  properties: ParagraphProperties;
  runs: Inline[];
  list?: { numId: string; level: number; marker: string; left?: number; hanging?: number };
}

export interface TableCell {
  blocks: Block[];
  gridSpan: number;
  rowSpan: number;
  width?: number;
  shading?: string;
  borders?: Record<string, string>;
  margins?: Partial<Record<"top" | "right" | "bottom" | "left", number>>;
  verticalAlign?: string;
}

export interface TableRow {
  cells: TableCell[];
  header?: boolean;
  height?: number;
  change?: ChangeMark;
}

export interface Table {
  type: "table";
  rows: TableRow[];
  widths: number[];
  width?: number;
  indent?: number;
  borders?: Record<string, string>;
  margins?: Partial<Record<"top" | "right" | "bottom" | "left", number>>;
}

export type Block = Paragraph | Table;

export interface SectionProperties {
  width: number;
  height: number;
  orientation: "portrait" | "landscape";
  margins: {
    top: number;
    right: number;
    bottom: number;
    left: number;
    header: number;
    footer: number;
  };
  titlePage?: boolean;
  breakType?: string;
  header?: Block[];
  firstHeader?: Block[];
  evenHeader?: Block[];
  footer?: Block[];
  firstFooter?: Block[];
  evenFooter?: Block[];
}

export interface DocxComment {
  id: string;
  author: string;
  initials?: string;
  date?: string;
  paragraphs: Block[];
  parentId?: string;
  done?: boolean;
  paraId?: string;
}

export interface DocxDocument {
  sections: Array<{ properties: SectionProperties; blocks: Block[] }>;
  comments: DocxComment[];
  evenAndOddHeaders?: boolean;
  imageParts: Record<string, { bytes: Uint8Array; mime: string }>;
}
