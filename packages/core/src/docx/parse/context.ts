import type { OpcPackage, Relationship, ThemeFonts } from "../../opc/package";
import type { Numbering } from "../numbering";
import type { StyleSheet } from "../styles";
import type { FieldStack } from "./fields";

export interface ParseContext {
  opc: OpcPackage;
  source: string;
  rels: Map<string, Relationship>;
  styles: StyleSheet;
  theme: ThemeFonts;
  numbering: Numbering;
  commentsOpen: Set<string>;
  images: Set<string>;
  fields: FieldStack;
  paragraphStyle?: string;
}
