import type { MediaPart } from "../common/media";
import type { Crop, Fill, LineStyle } from "../drawingml/fill";

export type { Crop, Fill, GradientStop, LineStyle } from "../drawingml/fill";

/** A rectangle in points (or in a group's child coordinate space). */
export interface Bounds {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

export interface Transform extends Bounds {
  /** Clockwise rotation in degrees. */
  readonly rotation: number;
  readonly flipH: boolean;
  readonly flipV: boolean;
}

export interface Geometry {
  /** DrawingML preset name (`rect`, `roundRect`, `ellipse`, ...) or `custom`. */
  readonly preset: string;
  /** Preset adjust values (`adj`, `adj1`, ...) in raw 1/100000 units. */
  readonly adjust: Readonly<Record<string, number>>;
  /** Custom geometry paths, each in its own coordinate space. */
  readonly paths?: readonly GeometryPath[];
}

export interface PathCommand {
  readonly type: "M" | "L" | "C" | "Q" | "A" | "Z";
  /** SVG path parameters in the path coordinate space. */
  readonly values: readonly number[];
}

export interface GeometryPath {
  readonly width: number;
  readonly height: number;
  readonly commands: readonly PathCommand[];
  readonly fill: boolean;
  readonly stroke: boolean;
}

export interface RunStyle {
  /** Font size in points. */
  readonly size?: number;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean;
  readonly strike?: boolean;
  readonly color?: string;
  readonly highlight?: string;
  readonly font?: string;
  readonly eastAsianFont?: string;
  /** Superscript/subscript offset as a fraction of the font size. */
  readonly baseline?: number;
  readonly caps?: "all" | "small";
  /** Letter spacing in points. */
  readonly spacing?: number;
}

export interface TextRun {
  readonly text: string;
  readonly style: RunStyle;
  readonly href?: string;
}

export type Spacing = { readonly points: number } | { readonly percent: number };

export interface Bullet {
  readonly text: string;
  readonly font?: string;
  readonly color?: string;
  /** Bullet size relative to the first run. */
  readonly scale?: number;
  /** Bullet size in points. */
  readonly size?: number;
}

export interface TextParagraph {
  readonly level: number;
  readonly align: "left" | "center" | "right" | "justify";
  /** Left margin in points. */
  readonly marginLeft: number;
  /** First-line indent in points (negative for hanging). */
  readonly indent: number;
  readonly lineSpacing?: Spacing;
  readonly spaceBefore?: Spacing;
  readonly spaceAfter?: Spacing;
  readonly bullet?: Bullet;
  readonly runs: readonly (TextRun | LineBreak)[];
  /** Style used to size empty paragraphs and line ends. */
  readonly endStyle: RunStyle;
}

export interface LineBreak {
  readonly type: "break";
  readonly style: RunStyle;
}

export interface Insets {
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
}

export interface TextBody {
  readonly anchor: "top" | "middle" | "bottom";
  readonly insets: Insets;
  readonly wrap: boolean;
  readonly vertical?: "vertical" | "vertical270" | "stacked";
  /** `normAutofit` font scale (1 = 100%). */
  readonly fontScale: number;
  /** `normAutofit` line spacing reduction (0–1). */
  readonly lineReduction: number;
  readonly paragraphs: readonly TextParagraph[];
}

interface ElementBase {
  readonly id: string;
  readonly name: string;
  readonly transform: Transform;
  readonly href?: string;
}

export interface ShapeElement extends ElementBase {
  readonly kind: "shape";
  readonly geometry: Geometry;
  readonly fill?: Fill;
  readonly line?: LineStyle;
  readonly text?: TextBody;
}

export interface PictureElement extends ElementBase {
  readonly kind: "picture";
  readonly geometry: Geometry;
  /** Media part path in {@link PptxPresentation.media}. */
  readonly part: string;
  readonly crop?: Crop;
  readonly line?: LineStyle;
  readonly description?: string;
}

export interface GroupElement extends ElementBase {
  readonly kind: "group";
  /** Coordinate space of the children. */
  readonly childBounds: Bounds;
  readonly children: readonly SlideElement[];
}

export interface CellBorders {
  readonly left?: LineStyle;
  readonly right?: LineStyle;
  readonly top?: LineStyle;
  readonly bottom?: LineStyle;
}

export interface TableCell {
  readonly text?: TextBody;
  readonly fill?: Fill;
  readonly borders: CellBorders;
  readonly columnSpan: number;
  readonly rowSpan: number;
  /** Covered by a neighbouring cell's span. */
  readonly merged: boolean;
}

export interface TableRow {
  /** Minimum row height in points. */
  readonly height: number;
  readonly cells: readonly TableCell[];
}

export interface TableElement extends ElementBase {
  readonly kind: "table";
  /** Column widths in points. */
  readonly columns: readonly number[];
  readonly rows: readonly TableRow[];
}

export type SlideElement = ShapeElement | PictureElement | GroupElement | TableElement;

export interface Slide {
  /** 1-based position in the presentation. */
  readonly number: number;
  readonly hidden: boolean;
  readonly background?: Fill;
  /** Master, layout and slide elements in paint order. */
  readonly elements: readonly SlideElement[];
}

export interface PptxPresentation {
  /** Slide width in points. */
  readonly width: number;
  /** Slide height in points. */
  readonly height: number;
  readonly slides: readonly Slide[];
  readonly media: Readonly<Record<string, MediaPart>>;
}
