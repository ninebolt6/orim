import { describe, expect, it } from "vite-plus/test";
import { attr, boolValue, children, descendants, elements, first, textValue, value } from "./xml";

describe("local-name XML helpers", () => {
  it("reads namespaced attributes and direct child elements", () => {
    const doc = new DOMParser().parseFromString(
      '<w:root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:item w:val="false"><w:text>hello</w:text></w:item><w:item w:val="1"/></w:root>',
      "application/xml",
    );
    const root = doc.documentElement!;
    const item = first(root, "item");
    expect(attr(item, "val")).toBe("false");
    expect(value(item)).toBe("false");
    expect(boolValue(item)).toBe(false);
    expect(textValue(first(item, "text"))).toBe("hello");
    expect(elements(root, "item")).toHaveLength(2);
    expect(children(root)).toHaveLength(2);
    expect(descendants(doc, "text")).toHaveLength(1);
  });

  it("ignores a matching local name in a foreign namespace", () => {
    const doc = new DOMParser().parseFromString(
      '<w:root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:foreign"><w:item x:val="wrong"/></w:root>',
      "application/xml",
    );
    expect(attr(first(doc.documentElement, "item"), "val")).toBeUndefined();
  });
});
