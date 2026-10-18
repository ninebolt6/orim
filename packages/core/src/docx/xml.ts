import { attr as commonAttr } from "../common/xml";

export { children, descendants, elements, first, R, textValue } from "../common/xml";

export const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
export const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
export const W15 = "http://schemas.microsoft.com/office/word/2012/wordml";
export const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
export const A = "http://schemas.openxmlformats.org/drawingml/2006/main";

export const attr = (
  element: Element | undefined,
  name: string,
  namespace: string | null = W,
): string | undefined => {
  return commonAttr(element, name, namespace);
};

export const value = (element: Element | undefined): string | undefined => {
  return attr(element, "val");
};

export const boolValue = (element: Element | undefined): boolean | undefined => {
  const item = value(element);
  if (item === undefined) return element ? true : undefined;
  return !["0", "false", "off", "no"].includes(item.toLowerCase());
};
