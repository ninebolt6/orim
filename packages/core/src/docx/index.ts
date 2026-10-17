export { parseDocx } from "./parse";
export { renderDocx, renderDocument } from "./render/document";
export { COMMENT_LANE_WIDTH_PT } from "./render/paginate";
export type {
  ChangesMode,
  RenderDocumentOptions,
  RenderOptions,
  RenderedDocument,
} from "./render/document";
export type {
  Block,
  ChangeKind,
  ChangeMark,
  DocxComment,
  DocxDocument,
  Inline,
  Paragraph,
  ParagraphProperties,
  RevisionInfo,
  SectionProperties,
  Table,
  TableCell,
  TableRow,
  TextStyle,
} from "./model";
