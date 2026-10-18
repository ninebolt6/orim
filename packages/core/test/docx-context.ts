import { zipSync } from "fflate";
import { openOpc } from "../src/opc/package";
import { parseNumbering } from "../src/docx/numbering";
import { parseStyles } from "../src/docx/styles";
import { createParseState, type ParseConfig, type ParseContext } from "../src/docx/parse/context";
import { xmlRoot } from "./xml";

export const parseContext = (overrides: Partial<ParseConfig> = {}): ParseContext => {
  return {
    config: {
      opc: openOpc(zipSync({})),
      source: "word/document.xml",
      rels: new Map(),
      styles: parseStyles(undefined),
      theme: {},
      numbering: parseNumbering(undefined),
      ...overrides,
    },
    state: createParseState(),
  };
};

export const wordElement = (xml: string): Element => {
  const root = xmlRoot(
    `<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${xml}</root>`,
  );
  return root.firstElementChild!;
};
