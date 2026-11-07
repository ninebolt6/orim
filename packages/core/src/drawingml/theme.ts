import { normalizeHex } from "../common/color";
import { attr, children, descendants, elements, first } from "../common/xml";

export interface ThemeFonts {
  readonly major?: string;
  readonly minor?: string;
  readonly majorEastAsia?: string;
  readonly minorEastAsia?: string;
}

export type ColorScheme = Readonly<Record<string, string>>;

/** Reads `a:clrScheme` slots (dk1, lt1, accent1, hlink, ...) as `#rrggbb`. */
export const themeColorScheme = (doc: Document | undefined): ColorScheme => {
  const scheme = doc && descendants(doc, "clrScheme")[0];
  if (!scheme) return {};
  return Object.fromEntries(
    children(scheme).flatMap((slot) => {
      const color = children(slot)[0];
      const value = normalizeHex(
        color?.localName === "sysClr"
          ? attr(color, "lastClr")
          : color?.localName === "srgbClr"
            ? attr(color, "val")
            : undefined,
      );
      return value ? [[slot.localName, value] as const] : [];
    }),
  );
};

export const themeFonts = (doc: Document | undefined): ThemeFonts => {
  if (!doc) return {};
  const major = descendants(doc, "majorFont")[0];
  const minor = descendants(doc, "minorFont")[0];
  const typeface = (font: Element | undefined, name: string) =>
    (font && attr(first(font, name), "typeface")) || undefined;
  return {
    major: typeface(major, "latin"),
    minor: typeface(minor, "latin"),
    majorEastAsia: typeface(major, "ea"),
    minorEastAsia: typeface(minor, "ea"),
  };
};

export interface ThemeFormats {
  readonly fills: readonly Element[];
  readonly lines: readonly Element[];
  readonly backgrounds: readonly Element[];
}

/** The `a:fmtScheme` style matrices that `fillRef`, `lnRef` and `bgRef` index into. */
export const themeFormats = (doc: Document | undefined): ThemeFormats => {
  const scheme = doc && descendants(doc, "fmtScheme")[0];
  const list = (name: string) => {
    const parent = first(scheme, name);
    return parent ? children(parent) : [];
  };
  return {
    fills: list("fillStyleLst"),
    lines: elements(first(scheme, "lnStyleLst"), "ln"),
    backgrounds: list("bgFillStyleLst"),
  };
};
