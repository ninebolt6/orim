import { descendants, elements } from "../xml";

export interface ThemeFonts {
  readonly major?: string;
  readonly minor?: string;
  readonly majorEastAsia?: string;
  readonly minorEastAsia?: string;
}

export function parseThemeFonts(doc: Document | undefined): ThemeFonts {
  if (!doc) return {};
  const major = descendants(doc, "majorFont")[0];
  const minor = descendants(doc, "minorFont")[0];
  const typeface = (font: Element | undefined, name: string) =>
    font ? (elements(font, name)[0]?.getAttribute("typeface") ?? undefined) : undefined;
  return {
    major: typeface(major, "latin"),
    minor: typeface(minor, "latin"),
    majorEastAsia: typeface(major, "ea"),
    minorEastAsia: typeface(minor, "ea"),
  };
}
