import type { SchemeResolver } from "../../drawingml/color";
import type { ThemeFonts, ThemeFormats } from "../../drawingml/theme";
import type { OpcPackage, Relationship } from "../../opc/package";
import type { ListStyle } from "./text";

/** A slide, layout or master part. */
export interface PartContext {
  readonly opc: OpcPackage;
  readonly path: string;
  readonly rels: ReadonlyMap<string, Relationship>;
}

export interface SlideContext {
  readonly scheme: SchemeResolver;
  readonly fonts: ThemeFonts;
  readonly formats: ThemeFormats;
  readonly slideNumber: number;
  /** Presentation `defaultTextStyle`, used by non-placeholder text. */
  readonly defaultText: ListStyle;
  readonly tableStyles: ReadonlyMap<string, Element>;
  readonly slidePaths: ReadonlyMap<string, number>;
}

/** Context for one shape: its part plus scheme overrides such as `phClr`. */
export interface ShapeContext extends SlideContext {
  readonly part: PartContext;
}

export const withPlaceholderColor = <T extends SlideContext>(
  context: T,
  color: string | undefined,
): T =>
  color
    ? {
        ...context,
        scheme: (name: string) => (name === "phClr" ? color : context.scheme(name)),
      }
    : context;

/** Drops `undefined` entries so spreading layers keeps inherited values. */
export const defined = <T extends object>(value: T): Partial<T> =>
  Object.fromEntries(Object.entries(value).filter(([, item]) => item !== undefined)) as Partial<T>;

export const layer = <T extends object>(
  ...layers: readonly (Partial<T> | undefined)[]
): Partial<T> =>
  layers.reduce<Partial<T>>(
    (result, item) => (item ? { ...result, ...defined(item) } : result),
    {},
  );
