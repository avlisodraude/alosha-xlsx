import { describe, expect, it } from "vitest";
import { resolveCellStyle, type Style } from "../../src/model/style.js";

describe("resolveCellStyle", () => {
  it("returns an empty style when nothing is set", () => {
    expect(resolveCellStyle()).toEqual({});
  });

  it("numFmt: cell overrides row overrides column", () => {
    const col: Style = { numFmt: "0.00" };
    const row: Style = { numFmt: "0%" };
    const cell: Style = { numFmt: "$#,##0" };

    expect(resolveCellStyle(row, col, cell).numFmt).toBe("$#,##0");
    expect(resolveCellStyle(row, col).numFmt).toBe("0%");
    expect(resolveCellStyle(undefined, col).numFmt).toBe("0.00");
  });

  it("font: row wins over column field-by-field", () => {
    const col: Style = { font: { name: "Arial", size: 10 } };
    const row: Style = { font: { bold: true, size: 12 } };

    expect(resolveCellStyle(row, col).font).toEqual({ name: "Arial", size: 12, bold: true });
  });

  it("partial font merge: cell setting only font.bold keeps row/col font fields", () => {
    const col: Style = { font: { name: "Arial" } };
    const row: Style = { font: { size: 14 } };
    const cell: Style = { font: { bold: true } };

    expect(resolveCellStyle(row, col, cell).font).toEqual({
      name: "Arial",
      size: 14,
      bold: true,
    });
  });

  it("alignment/protection/border merge the same way, row over column, cell over both", () => {
    const col: Style = {
      alignment: { horizontal: "left" },
      protection: { locked: true },
      border: { top: { style: "thin" } },
    };
    const row: Style = {
      alignment: { vertical: "middle" },
      protection: { hidden: false },
      border: { bottom: { style: "thin" } },
    };
    const cell: Style = {
      alignment: { wrapText: true },
      border: { top: { style: "thick" } },
    };

    const resolved = resolveCellStyle(row, col, cell);
    expect(resolved.alignment).toEqual({ horizontal: "left", vertical: "middle", wrapText: true });
    expect(resolved.protection).toEqual({ locked: true, hidden: false });
    expect(resolved.border).toEqual({
      top: { style: "thick" },
      bottom: { style: "thin" },
    });
  });

  it("fill is replaced wholesale (not merged) since it's a discriminated union", () => {
    const col: Style = { fill: { type: "pattern", pattern: "solid" } };
    const row: Style = { fill: { type: "pattern", pattern: "gray125" } };

    expect(resolveCellStyle(row, col).fill).toEqual({ type: "pattern", pattern: "gray125" });
    expect(resolveCellStyle(undefined, col).fill).toEqual({ type: "pattern", pattern: "solid" });
  });

  it("never mutates its inputs", () => {
    const col: Style = { font: { name: "Arial" } };
    const row: Style = { font: { bold: true } };
    const cell: Style = { font: { size: 11 } };
    const colSnapshot = structuredClone(col);
    const rowSnapshot = structuredClone(row);
    const cellSnapshot = structuredClone(cell);

    resolveCellStyle(row, col, cell);

    expect(col).toEqual(colSnapshot);
    expect(row).toEqual(rowSnapshot);
    expect(cell).toEqual(cellSnapshot);
  });

  it("omits a field entirely when none of row/col/cell set it", () => {
    const resolved = resolveCellStyle({ numFmt: "0" });
    expect(resolved.font).toBeUndefined();
    expect("font" in resolved).toBe(false);
    expect(resolved.fill).toBeUndefined();
    expect("fill" in resolved).toBe(false);
  });
});
