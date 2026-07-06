import ExcelJS, { Enums, ValueType, Workbook } from "@alosha/xlsx/compat";
import { describe, expect, it } from "vitest";

describe("@alosha/xlsx/compat", () => {
  it("resolves both the default ExcelJS-shaped namespace and named exports", () => {
    expect(ExcelJS.Workbook).toBe(Workbook);
    expect(ExcelJS.ValueType).toBe(ValueType);
    expect(Enums.ValueType).toBe(ValueType);
  });

  it("builds a workbook through the default export, matching ExcelJS usage", () => {
    const wb = new ExcelJS.Workbook();
    expect(wb).toBeInstanceOf(Workbook);
  });

  it("matches the ExcelJS ValueType numbering from COMPAT-SPEC.md §3", () => {
    expect(ValueType).toEqual({
      Null: 0,
      Merge: 1,
      Number: 2,
      String: 3,
      Date: 4,
      Hyperlink: 5,
      Formula: 6,
      SharedString: 7,
      RichText: 8,
      Boolean: 9,
      Error: 10,
    });
  });
});
