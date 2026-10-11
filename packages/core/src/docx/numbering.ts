import { attr, descendants, elements, first } from "../xml";

export interface NumberLevel {
  format: string;
  text: string;
  start: number;
  left?: number;
  hanging?: number;
  font?: string;
  suffix?: "tab" | "space" | "nothing";
}

export interface Numbering {
  marker(
    numId: string,
    level: number,
  ): {
    text: string;
    left?: number;
    hanging?: number;
    format: string;
    suffix?: "tab" | "space" | "nothing";
  };
}

function readLevel(node: Element): NumberLevel {
  const pPr = first(node, "pPr");
  const ind = first(pPr ?? node, "ind");
  const rFonts = first(first(node, "rPr") ?? node, "rFonts");
  const left = attr(ind, "left") ?? attr(ind, "start");
  const hanging = attr(ind, "hanging");
  const suffix = attr(first(node, "suff"), "val");
  return {
    format: attr(first(node, "numFmt"), "val") ?? "decimal",
    text: attr(first(node, "lvlText"), "val") ?? "%1.",
    start: Number(attr(first(node, "start"), "val") ?? 1),
    left: left === undefined ? undefined : Number(left) / 20,
    hanging: hanging === undefined ? undefined : Number(hanging) / 20,
    font: attr(rFonts, "ascii") ?? attr(rFonts, "hAnsi"),
    suffix: suffix === "space" || suffix === "nothing" || suffix === "tab" ? suffix : "tab",
  };
}

function alpha(number: number, upper: boolean): string {
  let result = "";
  for (let n = Math.max(1, number); n > 0; n = Math.floor((n - 1) / 26)) {
    result = String.fromCharCode((upper ? 65 : 97) + ((n - 1) % 26)) + result;
  }
  return result;
}

function roman(number: number): string {
  const digits: Array<[number, string]> = [
    [1000, "M"],
    [900, "CM"],
    [500, "D"],
    [400, "CD"],
    [100, "C"],
    [90, "XC"],
    [50, "L"],
    [40, "XL"],
    [10, "X"],
    [9, "IX"],
    [5, "V"],
    [4, "IV"],
    [1, "I"],
  ];
  return digits.reduce(
    (state, [value, symbol]) => ({
      remaining: state.remaining % value,
      output: state.output + symbol.repeat(Math.floor(state.remaining / value)),
    }),
    { remaining: number, output: "" },
  ).output;
}

function formatNumber(number: number, format: string): string {
  switch (format) {
    case "bullet":
      return "•";
    case "lowerLetter":
      return alpha(number, false);
    case "upperLetter":
      return alpha(number, true);
    case "lowerRoman":
      return roman(number).toLowerCase();
    case "upperRoman":
      return roman(number);
    case "decimalFullWidth":
      return String(number).replace(/\d/g, (digit) =>
        String.fromCharCode(digit.charCodeAt(0) + 0xfee0),
      );
    case "japaneseCounting":
      return String(number);
    default:
      return String(number);
  }
}

export function parseNumbering(doc: Document | undefined): Numbering {
  const abstract = new Map<string, Map<number, NumberLevel>>();
  const nums = new Map<string, string>();
  const overrides = new Map<string, Map<number, NumberLevel>>();
  if (doc) {
    for (const node of descendants(doc, "abstractNum")) {
      const id = attr(node, "abstractNumId");
      if (!id) continue;
      const levels = new Map<number, NumberLevel>();
      for (const level of elements(node, "lvl"))
        levels.set(Number(attr(level, "ilvl") ?? 0), readLevel(level));
      abstract.set(id, levels);
    }
    for (const node of descendants(doc, "num")) {
      const id = attr(node, "numId");
      const abstractId = attr(first(node, "abstractNumId"), "val");
      if (id && abstractId) {
        nums.set(id, abstractId);
        const levels = new Map<number, NumberLevel>();
        for (const override of elements(node, "lvlOverride")) {
          const index = Number(attr(override, "ilvl") ?? 0);
          const level = first(override, "lvl");
          const start = first(override, "startOverride");
          const value = level ? readLevel(level) : abstract.get(abstractId)?.get(index);
          if (value)
            levels.set(index, { ...value, start: Number(attr(start, "val") ?? value.start) });
        }
        overrides.set(id, levels);
      }
    }
  }
  const counts = new Map<string, number[]>();
  return {
    marker(numId, levelIndex) {
      const abstractId = nums.get(numId);
      const level = overrides.get(numId)?.get(levelIndex) ??
        abstract.get(abstractId ?? "")?.get(levelIndex) ?? {
          format: "decimal",
          text: `%${levelIndex + 1}.`,
          start: 1,
        };
      const values = counts.get(numId) ?? [];
      if (values[levelIndex] === undefined) values[levelIndex] = level.start - 1;
      values[levelIndex]++;
      values.length = levelIndex + 1;
      counts.set(numId, values);
      const label = level.text.replace(/%([1-9])/g, (_match, index: string) => {
        const which = Number(index) - 1;
        const fmt =
          overrides.get(numId)?.get(which)?.format ??
          abstract.get(abstractId ?? "")?.get(which)?.format ??
          level.format;
        return formatNumber(values[which] ?? 1, fmt);
      });
      const bullets: Record<string, string> = {
        "\uf0a7": "▪",
        "\uf0a8": "•",
        "\uf0b7": "•",
        "\uf0d8": "➢",
        "\uf076": "❖",
        "\uf0fc": "✓",
      };
      const marker =
        level.format === "bullet"
          ? label.replace(/[\ue000-\uf8ff]/gu, (char) => bullets[char] ?? "•")
          : label;
      return {
        text: marker,
        left: level.left,
        hanging: level.hanging,
        format: level.format,
        suffix: level.suffix,
      };
    },
  };
}
