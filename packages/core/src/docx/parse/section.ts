import { attr, first } from "../xml";
import type { SectionProperties } from "../model";
import { twips } from "../units";

export const sectionProperties = (node: Element | undefined): SectionProperties => {
  const pageSize = first(node, "pgSz");
  const pageMargins = first(node, "pgMar");
  const grid = first(node, "docGrid");
  const orientation = attr(pageSize, "orient") === "landscape" ? "landscape" : "portrait";
  const rawWidth = twips(attr(pageSize, "w"), 595.3);
  const rawHeight = twips(attr(pageSize, "h"), 841.9);
  const width = orientation === "landscape" ? Math.max(rawWidth, rawHeight) : rawWidth;
  const height = orientation === "landscape" ? Math.min(rawWidth, rawHeight) : rawHeight;
  return {
    width,
    height,
    orientation,
    margins: {
      top: twips(attr(pageMargins, "top"), 1440 / 20),
      right: twips(attr(pageMargins, "right"), 1440 / 20),
      bottom: twips(attr(pageMargins, "bottom"), 1440 / 20),
      left: twips(attr(pageMargins, "left"), 1440 / 20),
      header: twips(attr(pageMargins, "header"), 720 / 20),
      footer: twips(attr(pageMargins, "footer"), 720 / 20),
    },
    titlePage: !!first(node, "titlePg"),
    breakType: attr(first(node, "type"), "val") ?? "nextPage",
    linePitch:
      grid && ["lines", "linesAndChars"].includes(attr(grid, "type") ?? "")
        ? twips(attr(grid, "linePitch"), 0) || undefined
        : undefined,
  };
};
