import { attr } from "../xml";

export function twips(value: string | undefined, fallback: number): number;
export function twips(value: string | undefined): number | undefined;
export function twips(value: string | undefined, fallback?: number): number | undefined {
  const number = Number(value);
  return value !== undefined && Number.isFinite(number) ? number / 20 : fallback;
}

export function halfPoints(value: string | undefined): number | undefined {
  const number = Number(value);
  return value !== undefined && Number.isFinite(number) ? number / 2 : undefined;
}

export function points(value: string | undefined, divisor = 1): number | undefined {
  const number = Number(value);
  return value !== undefined && Number.isFinite(number) ? number / divisor : undefined;
}

export function cssWidth(element: Element | undefined): string | undefined {
  const value = attr(element, "w");
  if (value === undefined) return undefined;
  const type = attr(element, "type") ?? "dxa";
  if (type === "pct") {
    const percentage = value.endsWith("%") ? Number(value.slice(0, -1)) : Number(value) / 50;
    return Number.isFinite(percentage) ? `${percentage}%` : undefined;
  }
  if (type !== "dxa") return undefined;
  const twipValue = Number(value);
  return Number.isFinite(twipValue) ? `${twipValue / 20}pt` : undefined;
}
