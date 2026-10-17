import { describe, expect, it } from "vite-plus/test";
import { renderDocument } from "../src/docx/render/document";
import { renderWorksheet } from "../src/xlsx/render/worksheet";
import type { DocxDocument } from "../src/docx/model";
import type { XlsxWorkbook } from "../src/xlsx/model";

const docx: DocxDocument = {
  sections: [
    {
      properties: {
        width: 595,
        height: 842,
        orientation: "portrait",
        margins: { top: 48, right: 48, bottom: 48, left: 48, header: 20, footer: 20 },
      },
      blocks: [],
    },
  ],
  comments: [],
  imageParts: {},
};

const xlsx: XlsxWorkbook = {
  activeSheet: 0,
  date1904: false,
  sheets: [
    {
      name: "Sheet",
      state: "visible",
      rows: [],
      columns: [],
      merges: [],
      hyperlinks: [],
      defaultColumnWidth: 70,
      defaultRowHeight: 15,
      frozen: { rows: 0, columns: 0 },
      showGridLines: true,
      showHeaders: true,
    },
  ],
};

describe("cross-format render leases", () => {
  it("cancels a previous DOCX or XLSX render in the same container", () => {
    const container = document.createElement("div");
    const renderedDocx = renderDocument(docx, container);
    expect(container.querySelector(".orim-root")).not.toBeNull();
    const renderedXlsx = renderWorksheet(xlsx, 0, container);
    expect(container.querySelector(".orim-root")).toBeNull();
    expect(container.querySelector(".orim-xlsx")).not.toBeNull();
    const nextDocx = renderDocument(docx, container);
    expect(container.querySelector(".orim-xlsx")).toBeNull();
    expect(container.querySelector(".orim-root")).not.toBeNull();
    renderedDocx.destroy();
    renderedXlsx.destroy();
    nextDocx.destroy();
    expect(container.childElementCount).toBe(0);
  });
});
