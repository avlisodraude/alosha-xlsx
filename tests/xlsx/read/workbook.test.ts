import { describe, expect, it } from "vitest";
import { parseWorkbook } from "../../../src/xlsx/read/workbook.js";

function wb(inner: string): string {
  return `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${inner}</workbook>`;
}

describe("parseWorkbook", () => {
  it("reads the sheet list in document order with names, ids and rIds", () => {
    const header = parseWorkbook(
      wb(
        `<sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Second" sheetId="2" r:id="rId2"/></sheets>`,
      ),
    );
    expect(header.sheets).toEqual([
      { name: "Data", sheetId: 1, state: "visible", rId: "rId1" },
      { name: "Second", sheetId: 2, state: "visible", rId: "rId2" },
    ]);
  });

  it("defaults missing state to visible and reads hidden/veryHidden", () => {
    const header = parseWorkbook(
      wb(
        `<sheets><sheet name="A" sheetId="1" r:id="rId1"/><sheet name="B" sheetId="2" state="hidden" r:id="rId2"/><sheet name="C" sheetId="3" state="veryHidden" r:id="rId3"/></sheets>`,
      ),
    );
    expect(header.sheets.map((s) => s.state)).toEqual(["visible", "hidden", "veryHidden"]);
  });

  it("reads date1904 from workbookPr", () => {
    expect(parseWorkbook(wb(`<workbookPr date1904="1"/><sheets/>`)).date1904).toBe(true);
    expect(parseWorkbook(wb(`<sheets/>`)).date1904).toBe(false);
  });

  it("reads activeTab from bookViews, or leaves it undefined", () => {
    expect(
      parseWorkbook(wb(`<bookViews><workbookView activeTab="2"/></bookViews><sheets/>`)).activeTab,
    ).toBe(2);
    expect(
      parseWorkbook(wb(`<bookViews><workbookView/></bookViews><sheets/>`)).activeTab,
    ).toBeUndefined();
  });

  it("yields an empty header for a rootless document", () => {
    expect(parseWorkbook("<nope/>")).toEqual({ sheets: [], date1904: false });
  });
});
