import { beforeEach, describe, expect, it } from "vitest";
import { CellType, FormulaType } from "../../src/model/enums.js";
import { FakeWorksheetHost } from "./support/fake-host.js";

describe("Cell", () => {
  let host: FakeWorksheetHost;

  beforeEach(() => {
    host = new FakeWorksheetHost();
  });

  describe("value get/set for every type", () => {
    it("null", () => {
      const cell = host.getRow(1).getCell(1);
      cell.value = null;
      expect(cell.value).toBeNull();
      expect(cell.type).toBe(CellType.Null);
    });

    it("number", () => {
      const cell = host.getRow(1).getCell(1);
      cell.value = 42;
      expect(cell.value).toBe(42);
      expect(cell.type).toBe(CellType.Number);
    });

    it("string", () => {
      const cell = host.getRow(1).getCell(1);
      cell.value = "hello";
      expect(cell.value).toBe("hello");
      expect(cell.type).toBe(CellType.String);
    });

    it("boolean", () => {
      const cell = host.getRow(1).getCell(1);
      cell.value = true;
      expect(cell.value).toBe(true);
      expect(cell.type).toBe(CellType.Boolean);
    });

    it("date", () => {
      const cell = host.getRow(1).getCell(1);
      const d = new Date("2024-01-01T00:00:00.000Z");
      cell.value = d;
      expect(cell.value).toEqual(d);
      expect(cell.type).toBe(CellType.Date);
    });

    it("error", () => {
      const cell = host.getRow(1).getCell(1);
      cell.value = { error: "#REF!" };
      expect(cell.value).toEqual({ error: "#REF!" });
      expect(cell.type).toBe(CellType.Error);
    });

    it("richText", () => {
      const cell = host.getRow(1).getCell(1);
      cell.value = { richText: [{ text: "hi" }] };
      expect(cell.value).toEqual({ richText: [{ text: "hi" }] });
      expect(cell.type).toBe(CellType.RichText);
    });

    it("hyperlink", () => {
      const cell = host.getRow(1).getCell(1);
      cell.value = { text: "click", hyperlink: "https://example.com" };
      expect(cell.value).toEqual({ text: "click", hyperlink: "https://example.com" });
      expect(cell.type).toBe(CellType.Hyperlink);
    });

    it("formula", () => {
      const cell = host.getRow(1).getCell(1);
      cell.value = { formula: "SUM(A1:A2)", result: 3 };
      expect(cell.value).toEqual({ formula: "SUM(A1:A2)", result: 3 });
      expect(cell.type).toBe(CellType.Formula);
      expect(cell.formula).toBe("SUM(A1:A2)");
      expect(cell.result).toBe(3);
      expect(cell.formulaType).toBe(FormulaType.Master);
    });

    it("shared formula reports FormulaType.Shared", () => {
      const cell = host.getRow(1).getCell(1);
      cell.value = { sharedFormula: "A1", ref: "A1:A10" };
      expect(cell.formulaType).toBe(FormulaType.Shared);
    });
  });

  describe("merge", () => {
    it("setting a value on a merge slave writes the master", () => {
      const master = host.getRow(1).getCell(1);
      const slave = host.getRow(1).getCell(2);
      slave.merge(master);

      slave.value = "written through";
      expect(master.value).toBe("written through");
      expect(slave.value).toBe("written through");
    });

    it("reading a merge slave delegates to the master", () => {
      const master = host.getRow(1).getCell(1);
      master.value = "master value";
      const slave = host.getRow(1).getCell(2);
      slave.merge(master);

      expect(slave.value).toBe("master value");
      expect(slave.text).toBe("master value");
    });

    it("isMerged / master", () => {
      const master = host.getRow(1).getCell(1);
      const slave = host.getRow(1).getCell(2);

      expect(master.isMerged).toBe(false);
      expect(master.master).toBe(master);

      slave.merge(master);
      expect(slave.isMerged).toBe(true);
      expect(slave.master).toBe(master);
    });

    it("slave's own type is Merge even though value delegates", () => {
      const master = host.getRow(1).getCell(1);
      const slave = host.getRow(1).getCell(2);
      master.value = 1;
      slave.merge(master);

      expect(slave.type).toBe(CellType.Merge);
      expect(slave.value).toBe(1);
    });

    it("merge adopts the master's style unless ignoreStyle is set", () => {
      const master = host.getRow(1).getCell(1);
      master.font = { bold: true };

      const slave1 = host.getRow(1).getCell(2);
      slave1.merge(master);
      expect(slave1.style).toEqual(master.style);

      const slave2 = host.getRow(1).getCell(3);
      slave2.font = { italic: true };
      const slave2StyleBefore = slave2.style;
      slave2.merge(master, true);
      expect(slave2.style).toEqual(slave2StyleBefore);
    });

    it("unmerge restores independence and clears the value", () => {
      const master = host.getRow(1).getCell(1);
      master.value = "shared";
      const slave = host.getRow(1).getCell(2);
      slave.merge(master);

      slave.unmerge();
      expect(slave.isMerged).toBe(false);
      expect(slave.value).toBeNull();

      slave.value = "independent";
      expect(master.value).toBe("shared");
      expect(slave.value).toBe("independent");
    });
  });

  describe("text / address / row / col / $col$row", () => {
    it("address, row, col derive from the row/column that created the cell", () => {
      const cell = host.getRow(3).getCell(2);
      expect(cell.address).toBe("B3");
      expect(cell.row).toBe(3);
      expect(cell.col).toBe(2);
      expect(cell.$col$row).toBe("$B$3");
    });

    it("text / toString / toCsvString", () => {
      const cell = host.getRow(1).getCell(1);
      cell.value = "a,b";
      expect(cell.text).toBe("a,b");
      expect(cell.toString()).toBe("a,b");
      expect(cell.toCsvString()).toBe('"a,b"');
    });

    it("text renders a formula's cached result", () => {
      const cell = host.getRow(1).getCell(1);
      cell.value = { formula: "1+1", result: 2 };
      expect(cell.text).toBe("2");
    });
  });

  describe("style shortcuts", () => {
    it("write into cell.style", () => {
      const cell = host.getRow(1).getCell(1);
      cell.numFmt = "0.00";
      cell.font = { bold: true };
      cell.alignment = { horizontal: "center" };
      cell.border = { top: { style: "thin" } };
      cell.fill = { type: "pattern", pattern: "solid" };
      cell.protection = { locked: false };

      expect(cell.style).toEqual({
        numFmt: "0.00",
        font: { bold: true },
        alignment: { horizontal: "center" },
        border: { top: { style: "thin" } },
        fill: { type: "pattern", pattern: "solid" },
        protection: { locked: false },
      });
    });

    it("cell.style resolves row+column defaults at construction", () => {
      const row = host.getRow(5);
      row.font = { name: "Arial" };
      const column = host.getColumn(1);
      column.font = { bold: true };

      // A cell created after both row and column defaults are set should inherit the merge.
      const cell = row.getCell(1);
      expect(cell.font).toEqual({ name: "Arial", bold: true });
    });

    it("full style replace via the style setter", () => {
      const cell = host.getRow(1).getCell(1);
      cell.style = { numFmt: "0%" };
      expect(cell.style).toEqual({ numFmt: "0%" });
    });
  });

  describe("model round-trip", () => {
    it("round-trips a plain value", () => {
      const cell = host.getRow(1).getCell(1);
      cell.value = "hello";
      cell.font = { bold: true };
      const model = cell.model;

      const other = host.getRow(1).getCell(2);
      other.model = model;
      expect(other.value).toBe("hello");
      expect(other.font).toEqual({ bold: true });
      expect(other.model).toEqual({ ...model, address: other.address });
    });

    it("round-trips a formula", () => {
      const cell = host.getRow(1).getCell(1);
      cell.value = { formula: "A1+A2", result: 5 };
      const model = cell.model;
      expect(model).toEqual({
        address: "A1",
        type: CellType.Formula,
        formula: "A1+A2",
        result: 5,
      });

      const other = host.getRow(2).getCell(1);
      other.model = model;
      expect(other.formula).toBe("A1+A2");
      expect(other.result).toBe(5);
    });

    it("round-trips null with no value/style keys", () => {
      const cell = host.getRow(1).getCell(1);
      expect(cell.model).toEqual({ address: "A1", type: CellType.Null });
    });
  });

  describe("note / dataValidation clearing", () => {
    it("assigning undefined to note removes the comment", () => {
      const cell = host.getRow(1).getCell(1);
      cell.note = { texts: "hi" };
      expect(cell.note).toEqual({ texts: "hi" });
      cell.note = undefined;
      expect(cell.note).toBeUndefined();
      expect(host.comments.size).toBe(0);
    });

    it("assigning undefined to dataValidation removes the validation", () => {
      const cell = host.getRow(1).getCell(1);
      cell.dataValidation = { type: "whole", formulae: [1] };
      expect(cell.dataValidation).toEqual({ type: "whole", formulae: [1] });
      cell.dataValidation = undefined;
      expect(cell.dataValidation).toBeUndefined();
      expect(host.dataValidations.size).toBe(0);
    });

    it("clearing an absent note or validation is a no-op", () => {
      const cell = host.getRow(1).getCell(1);
      expect(() => {
        cell.note = undefined;
        cell.dataValidation = undefined;
      }).not.toThrow();
    });
  });
});
