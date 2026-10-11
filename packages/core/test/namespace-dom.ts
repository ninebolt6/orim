const getAttributeNS = Element.prototype.getAttributeNS;

const getNamespaceValue = (
  element: Element,
  namespace: string,
  localName: string,
): string | null => {
  for (const attribute of Array.from(element.attributes)) {
    const separator = attribute.name.indexOf(":");
    if (separator < 0 || attribute.name.slice(separator + 1) !== localName) continue;
    const prefix = attribute.name.slice(0, separator);
    let scope: Element | null = element;
    while (scope) {
      if (scope.getAttribute(`xmlns:${prefix}`) === namespace) return attribute.value;
      scope = scope.parentElement;
    }
  }
  return null;
};

Element.prototype.getAttributeNS = function (namespace, localName) {
  const value = getAttributeNS.call(this, namespace, localName);
  if (value !== null || !namespace) return value;
  return getNamespaceValue(this, namespace, localName);
};
