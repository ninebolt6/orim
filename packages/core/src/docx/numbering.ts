import { alphabetic, roman } from "../common/numbering";
import { attr, descendants, elements, first } from "./xml";
import type { Block } from "./model";

export interface NumberLevel {
  readonly format: string;
  readonly text: string;
  readonly start: number;
  readonly left?: number;
  readonly hanging?: number;
  readonly font?: string;
  readonly suffix?: "tab" | "space" | "nothing";
}

export interface NumberMarker {
  readonly text: string;
  readonly left?: number;
  readonly hanging?: number;
  readonly format: string;
  readonly suffix?: "tab" | "space" | "nothing";
  readonly counters: readonly number[];
}

export interface Numbering {
  level(numId: string, level: number): NumberLevel;
  marker(numId: string, level: number, counters?: readonly number[]): NumberMarker;
}

const integerValue = (value: unknown): number => {
  if (typeof value === "number") return value;
  if (typeof value === "string" && /^[+-]?\d+$/.test(value.trim())) return Number(value);
  return NaN;
};

export const safeListLevel = (value: unknown): number => {
  const level = integerValue(value);
  return Number.isInteger(level) ? Math.min(8, Math.max(0, level)) : 0;
};

export const safeListStart = (value: unknown): number => {
  const start = integerValue(value);
  return Number.isSafeInteger(start) ? start : 1;
};

const readLevel = (node: Element): NumberLevel => {
  const pPr = first(node, "pPr");
  const ind = first(pPr, "ind");
  const rFonts = first(first(node, "rPr"), "rFonts");
  const left = attr(ind, "left") ?? attr(ind, "start");
  const hanging = attr(ind, "hanging");
  const suffix = attr(first(node, "suff"), "val");
  return {
    format: attr(first(node, "numFmt"), "val") ?? "decimal",
    text: attr(first(node, "lvlText"), "val") ?? "%1.",
    start: safeListStart(attr(first(node, "start"), "val") ?? 1),
    left: left === undefined ? undefined : Number(left) / 20,
    hanging: hanging === undefined ? undefined : Number(hanging) / 20,
    font: attr(rFonts, "ascii") ?? attr(rFonts, "hAnsi"),
    suffix: suffix === "space" || suffix === "nothing" || suffix === "tab" ? suffix : "tab",
  };
};

const formatNumber = (number: number, format: string): string => {
  switch (format) {
    case "bullet":
      return "•";
    case "lowerLetter":
      return alphabetic(number, false);
    case "upperLetter":
      return alphabetic(number, true);
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
};

export const parseNumbering = (doc: Document | undefined): Numbering => {
  const abstract = new Map<string, Map<number, NumberLevel>>();
  const nums = new Map<string, string>();
  const overrides = new Map<string, Map<number, NumberLevel>>();
  if (doc) {
    for (const node of descendants(doc, "abstractNum")) {
      const id = attr(node, "abstractNumId");
      if (!id) continue;
      const levels = new Map<number, NumberLevel>();
      for (const level of elements(node, "lvl"))
        levels.set(safeListLevel(attr(level, "ilvl") ?? 0), readLevel(level));
      abstract.set(id, levels);
    }
    for (const node of descendants(doc, "num")) {
      const id = attr(node, "numId");
      const abstractId = attr(first(node, "abstractNumId"), "val");
      if (id && abstractId) {
        nums.set(id, abstractId);
        const levels = new Map<number, NumberLevel>();
        for (const override of elements(node, "lvlOverride")) {
          const index = safeListLevel(attr(override, "ilvl") ?? 0);
          const level = first(override, "lvl");
          const start = first(override, "startOverride");
          const value = level ? readLevel(level) : abstract.get(abstractId)?.get(index);
          if (value)
            levels.set(index, {
              ...value,
              start: safeListStart(attr(start, "val") ?? value.start),
            });
        }
        overrides.set(id, levels);
      }
    }
  }
  const level = (numId: string, levelIndex: number): NumberLevel => {
    const index = safeListLevel(levelIndex);
    const abstractId = nums.get(numId);
    return (
      overrides.get(numId)?.get(index) ??
      abstract.get(abstractId ?? "")?.get(index) ?? {
        format: "decimal",
        text: `%${index + 1}.`,
        start: 1,
      }
    );
  };
  return {
    level,
    marker(numId, levelIndex, counters = []) {
      const index = safeListLevel(levelIndex);
      const currentLevel = level(numId, index);
      const values = [...counters];
      if (values[index] === undefined) values[index] = currentLevel.start - 1;
      values[index]++;
      values.length = index + 1;
      const label = currentLevel.text.replace(/%([1-9])/g, (_match, index: string) => {
        const which = Number(index) - 1;
        const referencedLevel = level(numId, which);
        return formatNumber(values[which] ?? referencedLevel.start, referencedLevel.format);
      });
      const bullets: Readonly<Record<string, string>> = {
        "\uf0a7": "▪",
        "\uf0a8": "•",
        "\uf0b7": "•",
        "\uf0d8": "➢",
        "\uf076": "❖",
        "\uf0fc": "✓",
      };
      const marker =
        currentLevel.format === "bullet"
          ? label.replace(/[\ue000-\uf8ff]/gu, (char) => bullets[char] ?? "•")
          : label;
      return {
        text: marker,
        left: currentLevel.left,
        hanging: currentLevel.hanging,
        format: currentLevel.format,
        suffix: currentLevel.suffix,
        counters: values,
      };
    },
  };
};

export const assignListMarkers = (blocks: readonly Block[], numbering: Numbering): Block[] => {
  const counters = new Map<string, readonly number[]>();
  const assign = (block: Block): Block => {
    if (block.type === "paragraph") {
      if (!block.list) return block;
      const marker = numbering.marker(
        block.list.numId,
        block.list.level,
        counters.get(block.list.numId),
      );
      counters.set(block.list.numId, marker.counters);
      return { ...block, list: { ...block.list, marker: marker.text } };
    }
    return {
      ...block,
      rows: block.rows.map((row) => ({
        ...row,
        cells: row.cells.map((cell) => ({
          ...cell,
          blocks: cell.blocks.map(assign),
        })),
      })),
    };
  };
  return blocks.map(assign);
};
