import { attr, elements, first } from "../common/xml";
import { openOpc, type Relationship } from "../opc/package";
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

const relationship = (
  relationships: ReadonlyMap<string, Relationship>,
  suffix: string,
): Relationship | undefined => {
  return Array.from(relationships.values()).find(
    (item) => !item.external && item.type.endsWith(`/${suffix}`),
  );
};

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

export const parseXlsx = async (input: ArrayBuffer | Uint8Array | Blob): Promise<XlsxWorkbook> => {
  let bytes: Uint8Array;
  if (typeof Blob !== "undefined" && input instanceof Blob)
    bytes = new Uint8Array(await input.arrayBuffer());
  else if (input instanceof Uint8Array) bytes = input;
  else bytes = new Uint8Array(input as ArrayBuffer);
  const pkg = openOpc(bytes);
  const rootRelationships = await pkg.relationships("");
  const workbookRelationship = relationship(rootRelationships, "officeDocument");
  if (!workbookRelationship) throw new Error("Not an XLSX package: workbook part not found.");
  const workbookPath = pkg.resolve("", workbookRelationship.target);
  const workbookDocument = await pkg.xml(workbookPath);
  if (workbookDocument.documentElement.localName !== "workbook")
    throw new Error("Not an XLSX package: workbook part is not a workbook.");
  const workbookRelationships = await pkg.relationships(workbookPath);
  const stylesPath = relationship(workbookRelationships, "styles");
  const sharedPath = relationship(workbookRelationships, "sharedStrings");
  const themePath = relationship(workbookRelationships, "theme");
  const themeDocument = themePath
    ? await pkg.xml(pkg.resolve(workbookPath, themePath.target))
    : undefined;
  const theme = themeColors(themeDocument);
  const styleDocument = stylesPath
    ? await pkg.xml(pkg.resolve(workbookPath, stylesPath.target))
    : undefined;
  const styles = parseStyles(styleDocument, theme);
  const sharedDocument = sharedPath
    ? await pkg.xml(pkg.resolve(workbookPath, sharedPath.target))
    : undefined;
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
