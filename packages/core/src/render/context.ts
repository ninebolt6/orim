import type { DocxDocument } from "../docx/model";
import type { ChangesMode } from "./document";

export interface RenderContext {
  readonly doc: DocxDocument;
  readonly comments: boolean;
  readonly changes: ChangesMode;
  readonly objectUrls: Set<string>;
  readonly objectUrlsByPart: Map<string, string>;
  readonly colors: Map<string, string>;
}
