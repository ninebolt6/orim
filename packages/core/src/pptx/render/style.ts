import { validColor } from "../../common/sanitize";
import type { ObjectUrls } from "../../common/media";
import type { Fill, LineStyle, PptxPresentation } from "../model";

export interface RenderContext {
  readonly presentation: PptxPresentation;
  readonly urls: ObjectUrls;
  /** Returns a document-unique id for SVG definitions. */
  readonly id: () => string;
}

/** A length in slide points, scaled to the rendered slide width. */
export const pt = (value: number): string => `calc(var(--orim-pt) * ${Number(value.toFixed(3))})`;

export const fontFamily = (...fonts: readonly (string | undefined)[]): string =>
  [
    ...fonts.flatMap((font) => (font ? [`"${font.replace(/["\\]/g, "")}"`] : [])),
    "sans-serif",
  ].join(", ");

export const lineColor = (line: LineStyle | undefined): string | undefined =>
  line?.fill?.type === "solid"
    ? validColor(line.fill.color)
    : line?.fill?.type === "gradient"
      ? validColor(line.fill.stops[0]?.color)
      : undefined;

/** CSS `background` for slide backgrounds and table cells. */
export const fillBackground = (
  fill: Fill | undefined,
  context: RenderContext,
): string | undefined => {
  switch (fill?.type) {
    case "solid":
      return validColor(fill.color);
    case "gradient": {
      const stops = fill.stops
        .flatMap((stop) => {
          const color = validColor(stop.color);
          return color ? [`${color} ${Number((stop.position * 100).toFixed(2))}%`] : [];
        })
        .join(", ");
      if (!stops) return undefined;
      return fill.radial
        ? `radial-gradient(circle, ${stops})`
        : `linear-gradient(${Number((fill.angle + 90).toFixed(2))}deg, ${stops})`;
    }
    case "image": {
      const url = context.urls.url(fill.part, context.presentation.media[fill.part]);
      return url ? `center / 100% 100% no-repeat url("${url}")` : undefined;
    }
  }
  return undefined;
};

export const border = (line: LineStyle | undefined): string => {
  const color = lineColor(line);
  return color ? `${pt(line?.width ?? 0.75)} solid ${color}` : "none";
};
