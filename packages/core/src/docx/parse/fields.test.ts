import { describe, expect, it } from "vite-plus/test";
import { createFieldStack, fieldHref, fieldName, isPageField } from "./fields";

describe("complex field state", () => {
  it("keeps nested fields and instructions across paragraph boundaries", () => {
    const fields = createFieldStack();
    fields.begin();
    fields.append(" TOC ");
    fields.separate();
    fields.begin();
    fields.append(' HYPERLINK \\l "_TocOverview" ');
    expect(fields.separate()?.stage).toBe("result");
    expect(fieldName(fields.current()!.instruction)).toBe("HYPERLINK");
    expect(fieldHref(fields.current()!.instruction)).toBe("#_TocOverview");
    fields.end();
    expect(fields.some((field) => fieldName(field.instruction) === "TOC")).toBe(true);
    expect(isPageField(fields.current()!)).toBe(false);
    fields.end();
    fields.end();
  });

  it("recognizes PAGE and NUMPAGES as live fields", () => {
    const fields = createFieldStack();
    fields.begin();
    fields.append(" PAGE ");
    expect(isPageField(fields.current()!)).toBe(true);
    fields.end();
    fields.begin();
    fields.append(" NUMPAGES ");
    expect(isPageField(fields.current()!)).toBe(true);
  });
});
