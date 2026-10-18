import { describe, expect, it } from "vite-plus/test";
import { attr, children, descendants, elements, first, textValue } from "./xml";

describe("local-name XML helpers", () => {
  it("reads unprefixed attributes and direct child elements", () => {
    const doc = new DOMParser().parseFromString(
      '<root id="root"><item id="1"><text>hello</text></item><item id="2"/></root>',
      "application/xml",
    );
    const root = doc.documentElement!;
    const item = first(root, "item");
    expect(attr(root, "id")).toBe("root");
    expect(attr(item, "id")).toBe("1");
    expect(textValue(first(item, "text"))).toBe("hello");
    expect(elements(root, "item")).toHaveLength(2);
    expect(children(root)).toHaveLength(2);
    expect(descendants(doc, "text")).toHaveLength(1);
  });

  it("reads attributes in an explicitly selected namespace", () => {
    const doc = new DOMParser().parseFromString(
      '<w:root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:foreign"><w:item x:val="selected" w:val="ignored"/></w:root>',
      "application/xml",
    );
    const item = first(doc.documentElement, "item");
    expect(attr(item, "val", "urn:foreign")).toBe("selected");
    expect(attr(item, "val")).toBeUndefined();
  });
});
