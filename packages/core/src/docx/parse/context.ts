import type { OpcPackage, Relationship } from "../../opc/package";
import type { Numbering } from "../numbering";
import type { StyleSheet } from "../styles";
import type { ThemeFonts } from "../../drawingml/theme";
import { createFieldStack, type FieldStack } from "./fields";

export interface ParseConfig {
  readonly opc: OpcPackage;
  readonly source: string;
  readonly rels: ReadonlyMap<string, Relationship>;
  readonly styles: StyleSheet;
  readonly theme: ThemeFonts;
  readonly numbering: Numbering;
  readonly paragraphStyle?: string;
}

export interface ParseState {
  readonly fields: FieldStack;
  readonly commentRanges: CommentRangeState;
}

export interface ParseContext {
  readonly config: ParseConfig;
  readonly state: ParseState;
}

export interface CommentRangeState {
  start(id: string): void;
  end(id: string): void;
  snapshot(): string[];
}

export const createParseState = (): ParseState => {
  return {
    fields: createFieldStack(),
    commentRanges: createCommentRangeState(),
  };
};

export const createCommentRangeState = (): CommentRangeState => {
  const open = new Set<string>();
  return {
    start(id) {
      open.add(id);
    },
    end(id) {
      open.delete(id);
    },
    snapshot() {
      return [...open];
    },
  };
};
