import type { TextStyle } from "../model";

export interface FieldState {
  readonly instruction: string;
  readonly stage: "instruction" | "result";
  readonly style?: TextStyle;
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
}

export const createFieldStack = (): FieldStack => {
  const stack: MutableFieldState[] = [];
  const snapshot = (field: MutableFieldState): FieldState => ({ ...field });
  return {
    current() {
      const field = stack.at(-1);
      return field && snapshot(field);
    },
    some(predicate) {
      return stack.some((field) => predicate(snapshot(field)));
    },
    findLast(predicate) {
      return [...stack].reverse().map(snapshot).find(predicate);
    },
    begin(style) {
      stack.push({ instruction: "", stage: "instruction", style });
    },
    push(field) {
      stack.push({ ...field });
    },
    captureStyle(style) {
      for (const field of stack) {
        if (field.stage === "result" && !field.style && isPageField(field)) field.style = style;
      }
    },
    append(instruction) {
      const field = stack.at(-1);
      if (field?.stage === "instruction") field.instruction += instruction;
    },
    separate() {
      const field = stack.at(-1);
      if (field) field.stage = "result";
      return field && snapshot(field);
    },
    end() {
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

export const isPageField = (field: FieldState): boolean => {
  const name = fieldName(field.instruction);
  return name === "PAGE" || name === "NUMPAGES";
};
