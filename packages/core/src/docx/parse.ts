import type { PackageInput } from "../opc/package";
import type { DocxDocument } from "./model";
import { parseDocxModel } from "./parse/document";

export { parseDocxModel } from "./parse/document";

export const parseDocx = async (input: PackageInput): Promise<DocxDocument> => {
  return parseDocxModel(input);
};
