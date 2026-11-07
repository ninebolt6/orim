import { parseDocx, parsePptx, parseXlsx } from "@orim/core";
import type { DocxDocument, PptxPresentation, XlsxWorkbook } from "@orim/core";
import sampleDocxUrl from "../../../fixtures/sample.docx?url";
import samplePptxUrl from "../../../fixtures/sample.pptx?url";
import sampleXlsxUrl from "../../../fixtures/sample.xlsx?url";

export type View =
  | { readonly kind: "docx"; readonly document: DocxDocument }
  | { readonly kind: "xlsx"; readonly workbook: XlsxWorkbook }
  | { readonly kind: "pptx"; readonly presentation: PptxPresentation };

export type SampleKind = View["kind"];

const samples: Readonly<Record<SampleKind, string>> = {
  docx: sampleDocxUrl,
  xlsx: sampleXlsxUrl,
  pptx: samplePptxUrl,
};

export interface LoadResult {
  readonly view: View | null;
  readonly error: string;
}

export const parseFile = async (file: Blob, extension: string): Promise<View> => {
  if (extension === "docx") return { kind: "docx", document: await parseDocx(file) };
  if (extension === "xlsx") return { kind: "xlsx", workbook: await parseXlsx(file) };
  if (extension === "pptx") return { kind: "pptx", presentation: await parsePptx(file) };
  throw new Error(
    "Unsupported file format. Open a .docx, .xlsx or .pptx file. Save legacy .doc/.xls/.ppt files in one of these formats first.",
  );
};

export const loadSample = async (kind: SampleKind): Promise<View> => {
  const response = await fetch(samples[kind]);
  if (!response.ok) throw new Error(`Unable to load sample: ${response.status}`);
  return parseFile(await response.blob(), kind);
};
