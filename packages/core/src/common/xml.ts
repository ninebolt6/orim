export const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

export const attr = (
  element: Element | undefined,
  name: string,
  namespace: string | null = null,
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
