export const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
export const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
export const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
export const W15 = "http://schemas.microsoft.com/office/word/2012/wordml";
export const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
export const A = "http://schemas.openxmlformats.org/drawingml/2006/main";

export const attr = (
  element: Element | undefined,
  name: string,
  namespace: string | null = W,
): string | undefined => {
  if (!element) return undefined;
  const namespaced = element.getAttributeNS(namespace, name);
  if (namespaced !== null) return namespaced;
  const unprefixed = element.getAttributeNode(name);
  return unprefixed && unprefixed.namespaceURI === null ? unprefixed.value : undefined;
};

export const children = (parent: ParentNode): Element[] => {
  return Array.from(parent.childNodes).filter((node): node is Element => node.nodeType === 1);
};

export const elements = (parent: ParentNode | undefined, name: string): Element[] => {
  if (!parent) return [];
  return children(parent).filter((element) => element.localName === name);
};

export const descendants = (parent: ParentNode, name: string): Element[] => {
  const result: Element[] = [];
  const visit = (node: ParentNode): void => {
    for (const element of children(node)) {
      if (element.localName === name) result.push(element);
      visit(element);
    }
  };
  visit(parent);
  return result;
};

export const first = (parent: ParentNode | undefined, name: string): Element | undefined => {
  return elements(parent, name)[0];
};

export const textValue = (element: Element | undefined): string => {
  return element?.textContent ?? "";
};

export const value = (element: Element | undefined): string | undefined => {
  return attr(element, "val");
};

export const boolValue = (element: Element | undefined): boolean | undefined => {
  const item = value(element);
  if (item === undefined) return element ? true : undefined;
  return !["0", "false", "off", "no"].includes(item.toLowerCase());
};
