import { describe, expect, it } from "vite-plus/test";
import { attr, boolValue, value } from "./xml";

describe("DOCX XML helpers", () => {
  it("reads WordprocessingML attributes by default", () => {
    const doc = new DOMParser().parseFromString(
      '<w:root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:item w:val="false"/></w:root>',
      "application/xml",
    );
    const item = doc.documentElement!.firstElementChild!;
    expect(attr(item, "val")).toBe("false");
    expect(value(item)).toBe("false");
    expect(boolValue(item)).toBe(false);
  });

  it("ignores a matching local name in a foreign namespace", () => {
    const doc = new DOMParser().parseFromString(
      '<w:root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:foreign"><w:item x:val="wrong"/></w:root>',
      "application/xml",
    );
    expect(attr(doc.documentElement!.firstElementChild!, "val")).toBeUndefined();
  });
});
