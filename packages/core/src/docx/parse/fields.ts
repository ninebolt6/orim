import type { TextStyle } from "../model";

export interface FieldState {
  readonly instruction: string;
  readonly stage: "instruction" | "result";
  readonly style?: TextStyle;
  readonly name?: string;
  readonly href?: string;
}

export interface FieldStack {
  current(): FieldState | undefined;
  some(predicate: (field: FieldState) => boolean): boolean;
  findLast(predicate: (field: FieldState) => boolean): FieldState | undefined;
  begin(style?: TextStyle): void;
  push(field: FieldState): void;
  captureStyle(style: TextStyle): void;
  append(instruction: string): void;
  separate(): FieldState | undefined;
  end(): FieldState | undefined;
}

interface MutableFieldState {
  instruction: string;
  stage: "instruction" | "result";
  style?: TextStyle;
  name?: string;
  href?: string;
}

export const MAX_FIELD_DEPTH = 32;

export const createFieldStack = (): FieldStack => {
  const stack: MutableFieldState[] = [];
  let overflow = 0;
  const snapshot = (field: MutableFieldState): FieldState => ({ ...field });
  return {
    current() {
      const field = overflow ? undefined : stack.at(-1);
      return field && snapshot(field);
    },
    some(predicate) {
      return stack.some((field) => predicate(snapshot(field)));
    },
    findLast(predicate) {
      return [...stack].reverse().map(snapshot).find(predicate);
    },
    begin(style) {
      if (overflow || stack.length >= MAX_FIELD_DEPTH) overflow++;
      else stack.push({ instruction: "", stage: "instruction", style });
    },
    push(field) {
      if (overflow || stack.length >= MAX_FIELD_DEPTH) {
        overflow++;
        return;
      }
      const copy: MutableFieldState = { ...field };
      if (copy.stage === "result") {
        copy.name ??= fieldName(copy.instruction);
        if (!("href" in copy)) copy.href = fieldHref(copy.instruction);
      }
      stack.push(copy);
    },
    captureStyle(style) {
      for (const field of stack) {
        if (field.stage === "result" && !field.style && isPageField(field)) field.style = style;
      }
    },
    append(instruction) {
      if (overflow) return;
      const field = stack.at(-1);
      if (field?.stage === "instruction") {
        field.instruction += instruction;
        delete field.name;
        delete field.href;
      }
    },
    separate() {
      if (overflow) return undefined;
      const field = stack.at(-1);
      if (field) {
        field.stage = "result";
        field.name = fieldName(field.instruction);
        field.href = fieldHref(field.instruction);
      }
      return field && snapshot(field);
    },
    end() {
      if (overflow) {
        overflow--;
        return undefined;
      }
      const field = stack.pop();
      return field && snapshot(field);
    },
  };
};

export const fieldName = (instruction: string): string => {
  return instruction.trim().match(/^\S+/)?.[0]?.toUpperCase() ?? "";
};

export const fieldHref = (instruction: string): string | undefined => {
  if (fieldName(instruction) !== "HYPERLINK") return undefined;
  const bookmark = instruction.match(/\\l\s+(?:"([^"]+)"|(\S+))/i);
  if (bookmark) return `#${bookmark[1] ?? bookmark[2]}`;
  const target = instruction.match(/\bHYPERLINK\s+(?:"([^"]+)"|(\S+))/i);
  return target?.[1] ?? target?.[2];
};

export const fieldStateName = (field: FieldState): string => {
  return field.name ?? fieldName(field.instruction);
};

export const fieldStateHref = (field: FieldState): string | undefined => {
  return "href" in field ? field.href : fieldHref(field.instruction);
};

export const isPageField = (field: FieldState): boolean => {
  const name = fieldStateName(field);
  return name === "PAGE" || name === "NUMPAGES";
};
