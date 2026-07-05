import { beforeEach, describe, expect, it } from "vitest";
import { InvalidRowNumberError } from "../../src/errors.js";
import { FakeWorksheetHost } from "./support/fake-host.js";

describe("Row", () => {
  let host: FakeWorksheetHost;

  beforeEach(() => {
    host = new FakeWorksheetHost();
  });

  describe("getCell / findCell", () => {
    it("getCell creates on demand and is stable across calls", () => {
      const row = host.getRow(1);
      const cell = row.getCell(2);
      expect(row.getCell(2)).toBe(cell);
      expect(cell.address).toBe("B1");
    });

    it("getCell resolves a column letter", () => {
      const row = host.getRow(1);
      expect(row.getCell("C").col).toBe(3);
    });

    it("getCell resolves a column key ahead of letter parsing", () => {
      const row = host.getRow(1);
      const keyedColumn = host.getColumn(5);
      host.registerColumnKey("total", keyedColumn);
      expect(row.getCell("total").col).toBe(5);
    });

    it("findCell returns undefined for a column with no cell yet", () => {
      const row = host.getRow(1);
      expect(row.findCell(9)).toBeUndefined();
      row.getCell(9);
      expect(row.findCell(9)).toBeDefined();
    });
  });

  describe("values", () => {
    it("get returns a sparse array indexed by column number", () => {
      const row = host.getRow(1);
      row.getCell(2).value = "b";
      row.getCell(4).value = "d";

      const values = row.values as unknown[];
      expect(values[2]).toBe("b");
      expect(values[4]).toBe("d");
      expect(values[1]).toBeUndefined();
      expect(values[3]).toBeUndefined();
    });

    it("set from a contiguous array maps index 0 -> column 1", () => {
      const row = host.getRow(1);
      row.values = ["a", "b", "c"];
      expect(row.getCell(1).value).toBe("a");
      expect(row.getCell(2).value).toBe("b");
      expect(row.getCell(3).value).toBe("c");
    });

    it("set from a sparse array maps index -> column number directly", () => {
      const row = host.getRow(1);
      const sparse: string[] = [];
      sparse[5] = "e";
      sparse[10] = "j";
      row.values = sparse;

      expect(row.findCell(5)?.value).toBe("e");
      expect(row.findCell(10)?.value).toBe("j");
      expect(row.findCell(1)).toBeUndefined();
    });

    it("set from a keyed object maps through column keys", () => {
      const row = host.getRow(1);
      host.registerColumnKey("name", host.getColumn(1));
      host.registerColumnKey("age", host.getColumn(2));

      row.values = { name: "Ada", age: 36 };
      expect(row.getCell(1).value).toBe("Ada");
      expect(row.getCell(2).value).toBe(36);
    });

    it("round-trips get -> set without shifting columns", () => {
      const row = host.getRow(1);
      row.getCell(2).value = "b";
      row.getCell(4).value = "d";

      const other = host.getRow(2);
      other.values = row.values;
      expect(other.findCell(2)?.value).toBe("b");
      expect(other.findCell(4)?.value).toBe("d");
      expect(other.findCell(1)).toBeUndefined();
    });

    it("setter is destructive: clears prior cells", () => {
      const row = host.getRow(1);
      row.getCell(1).value = "old";
      row.values = ["new"];
      expect(row.getCell(1).value).toBe("new");
      expect(row.findCell(2)).toBeUndefined();
    });
  });

  describe("eachCell", () => {
    it("without includeEmpty visits only existing cells", () => {
      const row = host.getRow(1);
      row.getCell(2).value = "b";
      row.getCell(5).value = "e";

      const seen: number[] = [];
      row.eachCell((_cell, col) => seen.push(col));
      expect(seen).toEqual([2, 5]);
    });

    it("with includeEmpty visits every column up to the last populated one", () => {
      const row = host.getRow(1);
      row.getCell(3).value = "c";

      const seen: number[] = [];
      row.eachCell({ includeEmpty: true }, (_cell, col) => seen.push(col));
      expect(seen).toEqual([1, 2, 3]);
    });
  });

  describe("style fan-out", () => {
    it("row.font fans out to existing cells", () => {
      const row = host.getRow(1);
      row.getCell(1).value = "a";
      row.getCell(2).value = "b";

      row.font = { bold: true };
      expect(row.getCell(1).font).toEqual({ bold: true });
      expect(row.getCell(2).font).toEqual({ bold: true });
      expect(row.font).toEqual({ bold: true });
    });

    it("a cell created after the row style is set inherits it (via resolveCellStyle)", () => {
      const row = host.getRow(1);
      row.font = { bold: true };
      const cell = row.getCell(3);
      expect(cell.font).toEqual({ bold: true });
    });
  });

  describe("dimensions / hasValues / cellCount / actualCellCount", () => {
    it("are all empty for a fresh row", () => {
      const row = host.getRow(1);
      expect(row.dimensions).toBeNull();
      expect(row.hasValues).toBe(false);
      expect(row.cellCount).toBe(0);
      expect(row.actualCellCount).toBe(0);
    });

    it("reflect populated columns", () => {
      const row = host.getRow(1);
      row.getCell(2).value = "b";
      row.getCell(5).value = "e";

      expect(row.dimensions).toEqual({ min: 2, max: 5 });
      expect(row.cellCount).toBe(5);
      expect(row.actualCellCount).toBe(2);
      expect(row.hasValues).toBe(true);
    });

    it("hasValues is false when cells exist but are all Null", () => {
      const row = host.getRow(1);
      row.getCell(1); // touches/creates the cell without a value
      expect(row.actualCellCount).toBe(1);
      expect(row.hasValues).toBe(false);
    });
  });

  describe("splice", () => {
    it("removes cells and shifts the tail left, carrying value and style", () => {
      const row = host.getRow(1);
      row.getCell(1).value = "a";
      row.getCell(2).value = "b";
      row.getCell(3).value = "c";
      row.getCell(3).font = { bold: true };

      row.splice(2, 1);

      expect(row.getCell(1).value).toBe("a");
      expect(row.getCell(2).value).toBe("c");
      expect(row.getCell(2).font).toEqual({ bold: true });
      expect(row.findCell(3)).toBeUndefined();
    });

    it("inserts cells and shifts the tail right", () => {
      const row = host.getRow(1);
      row.getCell(1).value = "a";
      row.getCell(2).value = "b";

      row.splice(2, 0, "x", "y");

      expect(row.getCell(1).value).toBe("a");
      expect(row.getCell(2).value).toBe("x");
      expect(row.getCell(3).value).toBe("y");
      expect(row.getCell(4).value).toBe("b");
    });
  });

  describe("model round-trip", () => {
    it("round-trips number/cells/height/hidden/outlineLevel/style", () => {
      const row = host.getRow(1);
      row.getCell(1).value = "a";
      row.getCell(3).value = 42;
      row.height = 20;
      row.hidden = true;
      row.outlineLevel = 2;
      row.style = { numFmt: "0.00" };

      const model = row.model;
      const other = host.getRow(1);
      other.model = model;

      expect(other.getCell(1).value).toBe("a");
      expect(other.getCell(3).value).toBe(42);
      expect(other.height).toBe(20);
      expect(other.hidden).toBe(true);
      expect(other.outlineLevel).toBe(2);
      expect(other.style).toEqual({ numFmt: "0.00" });
    });

    it("throws InvalidRowNumberError when the model's number doesn't match", () => {
      const row = host.getRow(2);
      expect(() => {
        row.model = { number: 3, cells: [] };
      }).toThrow(InvalidRowNumberError);
    });
  });
});
