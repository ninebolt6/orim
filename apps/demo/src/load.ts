import { parseDocx, parseXlsx } from "@orim/core";
import type { DocxDocument, XlsxWorkbook } from "@orim/core";
import sampleDocxUrl from "../../../fixtures/sample.docx?url";
import sampleXlsxUrl from "../../../fixtures/sample.xlsx?url";

export type View =
  | { readonly kind: "docx"; readonly document: DocxDocument }
  | { readonly kind: "xlsx"; readonly workbook: XlsxWorkbook };

export interface LoadResult {
  readonly view: View | null;
  readonly error: string;
}

export const parseFile = async (file: Blob, extension: string): Promise<View> => {
  if (extension === "docx") return { kind: "docx", document: await parseDocx(file) };
  if (extension === "xlsx") return { kind: "xlsx", workbook: await parseXlsx(file) };
  throw new Error(
    "Unsupported file format. Open a .docx or .xlsx file. Save legacy .doc/.xls files in one of these formats first.",
  );
};

export const loadSample = async (kind: "docx" | "xlsx"): Promise<View> => {
  const response = await fetch(kind === "docx" ? sampleDocxUrl : sampleXlsxUrl);
  if (!response.ok) throw new Error(`Unable to load sample: ${response.status}`);
  return parseFile(await response.blob(), kind);
};
