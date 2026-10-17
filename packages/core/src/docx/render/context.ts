import type { DocxDocument } from "../model";
import type { ChangesMode } from "./document";

export interface PendingField {
  readonly element: HTMLElement;
  readonly name: string;
  readonly cached: string;
}

export interface RenderContext {
  readonly doc: DocxDocument;
  readonly comments: boolean;
  readonly changes: ChangesMode;
  readonly objectUrls: Set<string>;
  readonly objectUrlsByPart: Map<string, string>;
  readonly colors: Map<string, string>;
  readonly fields: PendingField[];
  readonly commentAnchors: Map<string, HTMLElement[]>;
  readonly commentBalloons: Map<string, HTMLElement>;
  readonly commentReferences: Map<string, number>;
  readonly decorationPageNumber?: number;
}
