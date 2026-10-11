import type { DocxDocument } from "../docx/model";
import type { ChangesMode } from "./document";

export interface RenderContext {
  doc: DocxDocument;
  comments: boolean;
  changes: ChangesMode;
  objectUrls: Set<string>;
  colors: Map<string, string>;
}
