import { normalizeHex } from "../common/color";
import { attr, children, descendants, elements, first } from "../common/xml";

export type EastAsianScript = "Jpan" | "Hang" | "Hans" | "Hant";

export interface ThemeFonts {
  readonly major?: string;
  readonly minor?: string;
  readonly majorEastAsia?: string;
  readonly minorEastAsia?: string;
  readonly majorEastAsiaScripts?: Readonly<Partial<Record<EastAsianScript, string>>>;
  readonly minorEastAsiaScripts?: Readonly<Partial<Record<EastAsianScript, string>>>;
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
  const scriptTypefaces = (font: Element | undefined) => {
    const scripts: Partial<Record<EastAsianScript, string>> = {};
    for (const item of elements(font, "font")) {
      const script = attr(item, "script");
      const value = attr(item, "typeface");
      if (
        value &&
        (script === "Jpan" || script === "Hang" || script === "Hans" || script === "Hant")
      )
        scripts[script] = value;
    }
    return Object.keys(scripts).length ? scripts : undefined;
  };
  return {
    major: typeface(major, "latin"),
    minor: typeface(minor, "latin"),
    majorEastAsia: typeface(major, "ea"),
    minorEastAsia: typeface(minor, "ea"),
    majorEastAsiaScripts: scriptTypefaces(major),
    minorEastAsiaScripts: scriptTypefaces(minor),
  };
};

const kana = /\p{Script=Hiragana}|\p{Script=Katakana}/u;
const hangul = /\p{Script=Hangul}/u;

/** Infers the East Asian script a run uses, preferring unambiguous text over language metadata. */
export const eastAsianScript = (
  language: string | undefined,
  text: string,
): EastAsianScript | undefined => {
  if (kana.test(text)) return "Jpan";
  if (hangul.test(text)) return "Hang";
  if (!language) return undefined;
  try {
    const script = new Intl.Locale(language).maximize().script;
    if (script === "Jpan") return "Jpan";
    if (script === "Kore") return "Hang";
    if (script === "Hans" || script === "Hant") return script;
  } catch {
    // Invalid language tags simply provide no script hint.
  }
  return undefined;
};

/** Resolves an empty `ea` theme slot through the script supplement matching the run. */
export const themeEastAsiaFont = (
  fonts: ThemeFonts,
  major: boolean,
  language: string | undefined,
  text: string,
): string | undefined => {
  const explicit = major ? fonts.majorEastAsia : fonts.minorEastAsia;
  if (explicit) return explicit;
  const script = eastAsianScript(language, text);
  const scripts = major ? fonts.majorEastAsiaScripts : fonts.minorEastAsiaScripts;
  return script ? scripts?.[script] : undefined;
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
