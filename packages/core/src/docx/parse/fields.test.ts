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
    expect(fieldName(fields.stack.at(-1)!.instruction)).toBe("HYPERLINK");
    expect(fieldHref(fields.stack.at(-1)!.instruction)).toBe("#_TocOverview");
    fields.end();
    expect(fields.stack).toHaveLength(1);
    expect(isPageField(fields.stack[0])).toBe(false);
    fields.end();
    fields.end();
  });

  it("recognizes PAGE and NUMPAGES as live fields", () => {
    const fields = createFieldStack();
    const page = fields.begin();
    fields.append(" PAGE ");
    expect(isPageField(page)).toBe(true);
    fields.end();
    const total = fields.begin();
    fields.append(" NUMPAGES ");
    expect(isPageField(total)).toBe(true);
  });
});
