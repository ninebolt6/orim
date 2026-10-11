import type { DocxDocument } from "./model";
import { parseDocxModel } from "./parse/document";

export { parseDocxModel } from "./parse/document";

export const parseDocx = async (input: ArrayBuffer | Uint8Array | Blob): Promise<DocxDocument> => {
  return parseDocxModel(input instanceof Blob ? await input.arrayBuffer() : input);
};
