import type { TextStyle } from "../model";

export interface FieldState {
  instruction: string;
  stage: "instruction" | "result";
  style?: TextStyle;
}

export interface FieldStack {
  readonly stack: FieldState[];
  begin(style?: TextStyle): FieldState;
  push(field: FieldState): void;
  append(instruction: string): void;
  separate(): FieldState | undefined;
  end(): FieldState | undefined;
}

export function createFieldStack(): FieldStack {
  const stack: FieldState[] = [];
  return {
    stack,
    begin(style) {
      const field = { instruction: "", stage: "instruction" as const, style };
      stack.push(field);
      return field;
    },
    push(field) {
      stack.push(field);
    },
    append(instruction) {
      const field = stack.at(-1);
      if (field?.stage === "instruction") field.instruction += instruction;
    },
    separate() {
      const field = stack.at(-1);
      if (field) field.stage = "result";
      return field;
    },
    end() {
      return stack.pop();
    },
  };
}

export function fieldName(instruction: string): string {
  return instruction.trim().match(/^\S+/)?.[0]?.toUpperCase() ?? "";
}

export function fieldHref(instruction: string): string | undefined {
  if (fieldName(instruction) !== "HYPERLINK") return undefined;
  const bookmark = instruction.match(/\\l\s+(?:"([^"]+)"|(\S+))/i);
  if (bookmark) return `#${bookmark[1] ?? bookmark[2]}`;
  const target = instruction.match(/\bHYPERLINK\s+(?:"([^"]+)"|(\S+))/i);
  return target?.[1] ?? target?.[2];
}

export function isPageField(field: FieldState): boolean {
  const name = fieldName(field.instruction);
  return name === "PAGE" || name === "NUMPAGES";
}
