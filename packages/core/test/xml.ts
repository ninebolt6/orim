export const xmlDocument = (source: string): Document => {
  return new DOMParser().parseFromString(source, "application/xml");
};

export const xmlRoot = (source: string): Element => {
  return xmlDocument(source).documentElement;
};
