import { describe, expect, it } from "vitest";
import { applyExcelJsStyle } from "../../src/compat/style-compat.js";
import { type Cell, resolveCellStyle, type Style, Workbook } from "../../src/index.js";

/** A fresh cell whose style is purely inherited (resolved from a row default at construction). */
function cellWithInheritedStyle(inherited: Style): Cell {
  const wb = new Workbook();
  const ws = wb.addWorksheet("S");
  const row = ws.getRow(1);
  row.style = inherited; // set before getCell so the cell resolves it at construction
  return row.getCell(1);
}

describe("applyExcelJsStyle — ExcelJS whole-object style precedence (COMPAT-SPEC.md §4)", () => {
  // The two behaviours side by side, on the SAME inputs, so the divergence is explicit.
  describe("the divergence: assigning a lone font over an inherited one", () => {
    const inheritedFont = { name: "Arial", size: 12 } as const;
    const patchFont = { bold: true } as const;

    it("core resolveCellStyle MERGES font fields (the baseline the shim diverges from)", () => {
      const merged = resolveCellStyle({ font: inheritedFont }, undefined, { font: patchFont });
      // bold coexists with the inherited name/size — deep field-level merge.
      expect(merged.font).toEqual({ name: "Arial", size: 12, bold: true });
    });

    it("applyExcelJsStyle REPLACES the whole font category (ExcelJS semantics)", () => {
      const cell = cellWithInheritedStyle({ font: inheritedFont });
      expect(cell.font).toEqual(inheritedFont); // inherited at construction

      applyExcelJsStyle(cell, { font: patchFont });

      // name/size are discarded — NOT merged. This is the whole point of the shim.
      expect(cell.font).toEqual({ bold: true });
      expect(cell.font).not.toHaveProperty("name");
      expect(cell.font).not.toHaveProperty("size");
    });
  });

  it("replaces each object category (fill/border/alignment/protection) outright", () => {
    const cell = cellWithInheritedStyle({
      fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF0000" } },
      border: { top: { style: "thin" }, bottom: { style: "thin" } },
      alignment: { horizontal: "center", vertical: "middle" },
      protection: { locked: true, hidden: true },
    });

    applyExcelJsStyle(cell, {
      fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FF00FF00" } },
      border: { left: { style: "thick" } },
      alignment: { wrapText: true },
      protection: { hidden: false },
    });

    expect(cell.fill).toEqual({ type: "pattern", pattern: "solid", bgColor: { argb: "FF00FF00" } });
    expect(cell.border).toEqual({ left: { style: "thick" } });
    expect(cell.alignment).toEqual({ wrapText: true });
    expect(cell.protection).toEqual({ hidden: false });
  });

  it("leaves categories the patch does not mention untouched (inherited values survive)", () => {
    const cell = cellWithInheritedStyle({
      font: { name: "Calibri", size: 11 },
      fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FF0000FF" } },
      numFmt: "0.00",
    });

    applyExcelJsStyle(cell, { font: { italic: true } });

    expect(cell.font).toEqual({ italic: true }); // replaced
    expect(cell.fill).toEqual({ type: "pattern", pattern: "solid", fgColor: { argb: "FF0000FF" } });
    expect(cell.numFmt).toBe("0.00"); // untouched
  });

  it("applies a multi-category patch, replacing each named category independently", () => {
    const cell = cellWithInheritedStyle({ font: { name: "Arial", size: 12 }, numFmt: "General" });

    applyExcelJsStyle(cell, { font: { bold: true }, numFmt: "0.0%" });

    expect(cell.font).toEqual({ bold: true });
    expect(cell.numFmt).toBe("0.0%");
  });

  it("sets a category the cell did not inherit", () => {
    const cell = cellWithInheritedStyle({ font: { name: "Arial" } });

    applyExcelJsStyle(cell, { fill: { type: "pattern", pattern: "gray125" } });

    expect(cell.fill).toEqual({ type: "pattern", pattern: "gray125" });
    expect(cell.font).toEqual({ name: "Arial" }); // untouched
  });

  it("clears a category when the patch sets it to undefined", () => {
    const cell = cellWithInheritedStyle({ font: { name: "Arial", size: 12 } });

    applyExcelJsStyle(cell, { font: undefined });

    expect(cell.font).toBeUndefined();
    expect(cell.style).not.toHaveProperty("font");
  });

  it("does not mutate the patch object", () => {
    const cell = cellWithInheritedStyle({ font: { name: "Arial" } });
    const patch: Style = { font: { bold: true } };
    const patchFontRef = patch.font;

    applyExcelJsStyle(cell, patch);

    expect(patch.font).toBe(patchFontRef);
    expect(patch).toEqual({ font: { bold: true } });
  });
});
