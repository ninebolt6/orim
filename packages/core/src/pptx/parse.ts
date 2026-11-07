import type { PackageInput } from "../opc/package";
import type { PptxPresentation } from "./model";
import { parsePptxModel } from "./parse/presentation";

export const parsePptx = (input: PackageInput): Promise<PptxPresentation> => parsePptxModel(input);
