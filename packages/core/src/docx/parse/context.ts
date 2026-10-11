import type { OpcPackage, Relationship, ThemeFonts } from "../../opc/package";
import type { Numbering } from "../numbering";
import type { StyleSheet } from "../styles";
import type { FieldStack } from "./fields";

export interface ParseContext {
  readonly opc: OpcPackage;
  readonly source: string;
  readonly rels: Map<string, Relationship>;
  readonly styles: StyleSheet;
  readonly theme: ThemeFonts;
  readonly numbering: Numbering;
  readonly commentRanges: CommentRangeState;
  readonly imageParts: ImagePathState;
  readonly fields: FieldStack;
  readonly paragraphStyle?: string;
}

export interface CommentRangeState {
  start(id: string): void;
  end(id: string): void;
  snapshot(): string[];
}

export interface ImagePathState {
  add(path: string): void;
  values(): readonly string[];
}

export function createCommentRangeState(): CommentRangeState {
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
}

export function createImagePathState(): ImagePathState {
  const paths = new Set<string>();
  return {
    add(path) {
      paths.add(path);
    },
    values() {
      return [...paths];
    },
  };
}
