import { attr } from "./xml";

const safeColor = (value: string | undefined): string | undefined => {
  return value && /^[\da-f]{6}$/i.test(value) ? `#${value}` : undefined;
};

export const border = (element: Element | undefined): string | undefined => {
  if (!element) return undefined;
  const value = attr(element, "val");
  if (value === "nil" || value === "none") return "none";
  const requestedSize = Number(attr(element, "sz") ?? 4) / 8;
  const size = Math.max(0.25, Number.isFinite(requestedSize) ? requestedSize : 0.5);
  const style =
    value === "double"
      ? "double"
      : value === "dotted"
        ? "dotted"
        : value === "dashed" || value === "dashSmallGap"
          ? "dashed"
          : "solid";
  const color = safeColor(attr(element, "color")) ?? "currentColor";
  return `${size}pt ${style} ${color}`;
};
