export function attr(element: Element | undefined, name: string): string | undefined {
  if (!element) return undefined;
  for (const item of Array.from(element.attributes)) {
    if (item.localName === name || item.name.split(":").at(-1) === name) return item.value;
  }
  return undefined;
}

export function children(parent: ParentNode): Element[] {
  return Array.from(parent.childNodes).filter((node): node is Element => node.nodeType === 1);
}

export function elements(parent: ParentNode | undefined, name: string): Element[] {
  if (!parent) return [];
  return children(parent).filter((element) => element.localName === name);
}

export function descendants(parent: ParentNode, name: string): Element[] {
  const result: Element[] = [];
  const visit = (node: ParentNode): void => {
    for (const element of children(node)) {
      if (element.localName === name) result.push(element);
      visit(element);
    }
  };
  visit(parent);
  return result;
}

export function first(parent: ParentNode | undefined, name: string): Element | undefined {
  return elements(parent, name)[0];
}

export function textValue(element: Element | undefined): string {
  return element?.textContent ?? "";
}

export function value(element: Element | undefined): string | undefined {
  return attr(element, "val");
}

export function boolValue(element: Element | undefined): boolean | undefined {
  const item = value(element);
  if (item === undefined) return element ? true : undefined;
  return !["0", "false", "off", "no"].includes(item.toLowerCase());
}
