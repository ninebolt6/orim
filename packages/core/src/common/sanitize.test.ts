import { describe, expect, it } from "vite-plus/test";
import { safeHref, validColor } from "./sanitize";

describe("render input sanitization", () => {
  it("accepts safe links and rejects executable schemes", () => {
    expect(safeHref("#_TocOverview")).toBe("#_TocOverview");
    expect(safeHref("https://example.com/")).toBe("https://example.com/");
    expect(safeHref("mailto:reader@example.com")).toBe("mailto:reader@example.com");
    expect(safeHref("javascript:alert(1)")).toBeUndefined();
    expect(safeHref("data:text/html,bad")).toBeUndefined();
    expect(safeHref("../document.xml")).toBeUndefined();
  });

  it("accepts only hexadecimal colors", () => {
    expect(validColor("#aabbcc")).toBe("#aabbcc");
    expect(validColor("red;background:url(x)")).toBeUndefined();
  });
});
