import { attr, elements, first } from "../common/xml";
import { loadOpc, relatedPart, type PackageInput } from "../opc/package";
import type { Worksheet, XlsxWorkbook } from "./model";
import { themeColors } from "./colors";
import { defaultCellStyle, parseStyles } from "./parse/styles";
import {
  boolAttribute,
  integerAttribute,
  parseSharedStrings,
  parseWorksheet,
  relationshipId,
} from "./parse/worksheet";

export const activeSheetIndex = (
  sheets: readonly Worksheet[],
  activeTab: number | undefined,
  sheetIndexes: ReadonlyMap<number, number>,
): number => {
  const mappedActive = activeTab === undefined ? undefined : sheetIndexes.get(activeTab);
  const firstVisible = sheets.findIndex((sheet) => sheet.state === "visible");
  return mappedActive !== undefined && sheets[mappedActive]!.state === "visible"
    ? mappedActive
    : firstVisible >= 0
      ? firstVisible
      : 0;
};

export const parseXlsx = async (input: PackageInput): Promise<XlsxWorkbook> => {
  const pkg = await loadOpc(input);
  const workbookPath = await relatedPart(pkg, "", "officeDocument");
  if (!workbookPath) throw new Error("Not an XLSX package: workbook part not found.");
  const workbookDocument = await pkg.xml(workbookPath);
  if (workbookDocument.documentElement.localName !== "workbook")
    throw new Error("Not an XLSX package: workbook part is not a workbook.");
  const workbookRelationships = await pkg.relationships(workbookPath);
  const relatedXml = async (type: string): Promise<Document | undefined> => {
    const path = await relatedPart(pkg, workbookPath, type);
    return path && pkg.zip.has(path) ? pkg.xml(path) : undefined;
  };
  const theme = themeColors(await relatedXml("theme"));
  const styles = parseStyles(await relatedXml("styles"), theme);
  const sharedDocument = await relatedXml("sharedStrings");
  const shared = parseSharedStrings(sharedDocument?.documentElement, styles);
  const workbookRoot = workbookDocument.documentElement;
  const workbookProperties = first(workbookRoot, "workbookPr");
  const date1904 = boolAttribute(workbookProperties, "date1904", false);
  const sheetEntries = elements(first(workbookRoot, "sheets"), "sheet");
  const parsedSheets = await Promise.all(
    sheetEntries.map(async (sheet) => {
      const id = relationshipId(sheet);
      const rel = id ? workbookRelationships.get(id) : undefined;
      if (!rel || rel.external || !rel.type.endsWith("/worksheet")) return undefined;
      const state = attr(sheet, "state");
      const sheetPath = pkg.resolve(workbookPath, rel.target);
      const document = await pkg.xml(sheetPath);
      const relationships = await pkg.relationships(sheetPath);
      return parseWorksheet(document.documentElement, {
        name: attr(sheet, "name") ?? "",
        state: state === "hidden" || state === "veryHidden" ? state : "visible",
        date1904,
        shared,
        styles,
        relationships,
      });
    }),
  );
  const sheets: Worksheet[] = [];
  const sheetIndexes = new Map<number, number>();
  parsedSheets.forEach((sheet, index) => {
    if (sheet) {
      sheetIndexes.set(index, sheets.length);
      sheets.push(sheet);
    }
  });
  const workbookView = elements(first(workbookRoot, "bookViews"), "workbookView")[0];
  const activeTab = integerAttribute(workbookView, "activeTab");
  const activeSheet = activeSheetIndex(sheets, activeTab, sheetIndexes);
  return { sheets, activeSheet, date1904 };
};

export { defaultCellStyle };
