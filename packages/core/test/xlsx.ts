import { strToU8, zipSync } from "fflate";

export interface XlsxPart {
  readonly path?: string;
  readonly xml: string;
}

export interface XlsxRelationship {
  readonly id: string;
  readonly type: string;
  readonly target: string;
  readonly external?: boolean;
}

export interface XlsxSheet {
  readonly name: string;
  readonly xml: string;
  readonly path?: string;
  readonly state?: "visible" | "hidden" | "veryHidden";
  readonly kind?: "worksheet" | "chartsheet";
  readonly relationships?: readonly XlsxRelationship[];
}

export interface BuildXlsxOptions {
  readonly workbookPath?: string;
  readonly sheets?: readonly XlsxSheet[];
  readonly sharedStrings?: XlsxPart;
  readonly styles?: XlsxPart;
  readonly theme?: XlsxPart;
  readonly date1904?: boolean;
  readonly activeTab?: number;
  readonly includeWorkbook?: boolean;
  readonly strict?: boolean;
}

const relativeTarget = (source: string, target: string): string => {
  const sourceParts = source.split("/").slice(0, -1);
  const targetParts = target.split("/");
  let common = 0;
  while (
    common < sourceParts.length &&
    common < targetParts.length &&
    sourceParts[common] === targetParts[common]
  )
    common++;
  return [...sourceParts.slice(common).map(() => ".."), ...targetParts.slice(common)].join("/");
};

const relsPath = (path: string): string => {
  const parts = path.split("/");
  const file = parts.pop() ?? "";
  return [...parts, "_rels", `${file}.rels`].filter(Boolean).join("/");
};

const workbookRelationship = (
  id: string,
  type: string,
  target: string,
  namespace: string,
): string => {
  return `<Relationship Id="${id}" Type="${namespace}${type}" Target="${target}"/>`;
};

export const buildXlsx = (options: BuildXlsxOptions = {}): Uint8Array => {
  const workbookPath = options.workbookPath ?? "custom/package/workbook.xml";
  const sheets = options.sheets ?? [];
  const partTypes = new Map<string, string>();
  const archive: Record<string, Uint8Array> = {};
  const workbookRelationships: string[] = [];
  const workbookSheets: string[] = [];
  const relationshipNamespace = options.strict
    ? "http://purl.oclc.org/ooxml/officeDocument/relationships/"
    : "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
  const registerPart = (path: string, xml: string, contentType: string): void => {
    archive[path] = strToU8(xml);
    partTypes.set(path, contentType);
  };

  sheets.forEach((sheet, index) => {
    const path = sheet.path ?? `content/grids/sheet${index + 1}.xml`;
    const id = `rSheet${index + 1}`;
    const kind = sheet.kind ?? "worksheet";
    const relationshipTarget = relativeTarget(workbookPath, path);
    workbookRelationships.push(
      workbookRelationship(id, kind, relationshipTarget, relationshipNamespace),
    );
    workbookSheets.push(
      `<sheet name="${sheet.name}"${sheet.state && sheet.state !== "visible" ? ` state="${sheet.state}"` : ""} r:id="${id}"/>`,
    );
    registerPart(
      path,
      sheet.xml,
      kind === "worksheet"
        ? "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
        : "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml",
    );
    if (sheet.relationships?.length) {
      const relationships = sheet.relationships
        .map(
          (item) =>
            `<Relationship Id="${item.id}" Type="${item.type}" Target="${item.target}"${item.external ? ' TargetMode="External"' : ""}/>`,
        )
        .join("");
      archive[relsPath(path)] = strToU8(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships}</Relationships>`,
      );
    }
  });

  const optionalParts = [
    options.sharedStrings && {
      part: options.sharedStrings,
      id: "rShared",
      type: "sharedStrings",
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
    },
    options.styles && {
      part: options.styles,
      id: "rStyles",
      type: "styles",
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
    },
    options.theme && {
      part: options.theme,
      id: "rTheme",
      type: "theme",
      contentType: "application/vnd.openxmlformats-officedocument.theme+xml",
    },
  ].filter((item): item is NonNullable<typeof item> => Boolean(item));
  for (const item of optionalParts) {
    const path = item.part.path ?? `custom/parts/${item.type}.xml`;
    workbookRelationships.push(
      workbookRelationship(
        item.id,
        item.type,
        relativeTarget(workbookPath, path),
        relationshipNamespace,
      ),
    );
    registerPart(path, item.part.xml, item.contentType);
  }

  if (options.includeWorkbook !== false) {
    const workbookXml =
      `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `<workbookPr${options.date1904 ? ' date1904="1"' : ""}/>` +
      `<bookViews><workbookView activeTab="${options.activeTab ?? 0}"/></bookViews>` +
      `<sheets>${workbookSheets.join("")}</sheets></workbook>`;
    registerPart(
      workbookPath,
      workbookXml,
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    );
    archive["_rels/.rels"] = strToU8(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${workbookRelationship(
        "rWorkbook",
        "officeDocument",
        workbookPath,
        relationshipNamespace,
      )}</Relationships>`,
    );
    archive[relsPath(workbookPath)] = strToU8(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${workbookRelationships.join("")}</Relationships>`,
    );
  }

  archive["[Content_Types].xml"] = strToU8(
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>${Array.from(
      partTypes,
      ([path, contentType]) => `<Override PartName="/${path}" ContentType="${contentType}"/>`,
    ).join("")}</Types>`,
  );
  return zipSync(archive, { level: 0 });
};
