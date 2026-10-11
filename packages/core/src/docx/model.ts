export type ChangeKind = "ins" | "del" | "moveFrom" | "moveTo";

export interface RevisionInfo {
  readonly author: string;
  readonly date: string;
}

export type ChangeMark = RevisionInfo & {
  readonly kind: ChangeKind;
};

export interface TextStyle {
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: string;
  readonly strike?: boolean;
  readonly doubleStrike?: boolean;
  readonly color?: string;
  readonly highlight?: string;
  readonly fontFamily?: string;
  readonly fontSize?: number;
  readonly verticalAlign?: "superscript" | "subscript";
  readonly caps?: boolean;
  readonly smallCaps?: boolean;
  readonly hidden?: boolean;
  readonly backgroundColor?: string;
}

export type Inline =
  | {
      readonly type: "text";
      readonly text: string;
      readonly style?: TextStyle;
      readonly change?: ChangeMark;
      readonly formattingChange?: RevisionInfo;
      readonly comments?: readonly string[];
      readonly href?: string;
    }
  | {
      readonly type: "image";
      readonly path: string;
      readonly width?: number;
      readonly height?: number;
      readonly change?: ChangeMark;
    }
  | {
      readonly type: "field";
      readonly name: string;
      readonly cached?: string;
      readonly style?: TextStyle;
    }
  | { readonly type: "break"; readonly page?: boolean }
  | { readonly type: "bookmark"; readonly name: string }
  | { readonly type: "comment-reference"; readonly id: string };

export interface ParagraphProperties {
  readonly align?: string;
  readonly styleId?: string;
  readonly before?: number;
  readonly after?: number;
  readonly line?: number;
  readonly lineRule?: string;
  readonly left?: number;
  readonly right?: number;
  readonly firstLine?: number;
  readonly hanging?: number;
  readonly contextualSpacing?: boolean;
  readonly pageBreakBefore?: boolean;
  readonly pageBreakAfter?: boolean;
  readonly keepNext?: boolean;
  readonly snapToGrid?: boolean;
  readonly shading?: string;
  readonly borders?: Readonly<Partial<Record<"top" | "right" | "bottom" | "left", string>>>;
  readonly borderSpacing?: Readonly<Partial<Record<"top" | "right" | "bottom" | "left", number>>>;
  readonly change?: ChangeMark;
  readonly formattingChange?: RevisionInfo;
}

export interface ListProperties {
  readonly numId: string;
  readonly level: number;
  readonly marker?: string;
  readonly left?: number;
  readonly hanging?: number;
  readonly suffix?: "tab" | "space" | "nothing";
}

export interface Paragraph {
  readonly type: "paragraph";
  readonly properties: ParagraphProperties;
  readonly runs: readonly Inline[];
  readonly markStyle?: TextStyle;
  readonly list?: ListProperties;
}

export interface TableCell {
  readonly blocks: readonly Block[];
  readonly gridSpan: number;
  readonly rowSpan: number;
  readonly width?: string;
  readonly shading?: string;
  readonly borders?: Readonly<Record<string, string>>;
  readonly margins?: Readonly<Partial<Record<"top" | "right" | "bottom" | "left", number>>>;
  readonly verticalAlign?: string;
}

export interface TableRow {
  readonly cells: readonly TableCell[];
  readonly header?: boolean;
  readonly height?: number;
  readonly change?: ChangeMark;
}

export interface Table {
  readonly type: "table";
  readonly rows: readonly TableRow[];
  readonly widths: readonly (string | undefined)[];
  readonly width?: string;
  readonly indent?: number;
  readonly align?: string;
  readonly borders?: Readonly<Record<string, string>>;
  readonly margins?: Readonly<Partial<Record<"top" | "right" | "bottom" | "left", number>>>;
}

export type Block = Paragraph | Table;

export interface SectionProperties {
  readonly width: number;
  readonly height: number;
  readonly orientation: "portrait" | "landscape";
  readonly margins: {
    readonly top: number;
    readonly right: number;
    readonly bottom: number;
    readonly left: number;
    readonly header: number;
    readonly footer: number;
  };
  readonly titlePage?: boolean;
  readonly breakType?: string;
  readonly linePitch?: number;
  readonly header?: readonly Block[];
  readonly firstHeader?: readonly Block[];
  readonly evenHeader?: readonly Block[];
  readonly footer?: readonly Block[];
  readonly firstFooter?: readonly Block[];
  readonly evenFooter?: readonly Block[];
}

export interface DocxComment {
  readonly id: string;
  readonly author: string;
  readonly initials?: string;
  readonly date?: string;
  readonly paragraphs: readonly Block[];
  readonly parentId?: string;
  readonly done?: boolean;
  readonly paraId?: string;
}

export interface DocxSection {
  readonly properties: SectionProperties;
  readonly blocks: readonly Block[];
}

export interface ReadonlyBytes {
  readonly length: number;
  readonly [index: number]: number;
  slice(start?: number, end?: number): Uint8Array;
}

export interface ImagePart {
  readonly bytes: ReadonlyBytes;
  readonly mime: string;
}

export interface DocxDocument {
  readonly sections: readonly DocxSection[];
  readonly comments: readonly DocxComment[];
  readonly evenAndOddHeaders?: boolean;
  readonly defaultTabStop?: number;
  readonly imageParts: Readonly<Record<string, ImagePart>>;
}
