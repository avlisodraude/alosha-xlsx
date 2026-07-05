import { beforeEach, describe, expect, it } from "vitest";
import { Column, type ColumnRange, DEFAULT_COLUMN_WIDTH } from "../../src/model/column.js";
import { FakeWorksheetHost } from "./support/fake-host.js";

describe("Column", () => {
  let host: FakeWorksheetHost;

  beforeEach(() => {
    host = new FakeWorksheetHost();
  });

  it("number / letter", () => {
    const col = host.getColumn(28);
    expect(col.number).toBe(28);
    expect(col.letter).toBe("AB");
  });

  describe("header writes top cells", () => {
    it("a single string writes row 1", () => {
      const col = host.getColumn(2);
      col.header = "Name";
      expect(host.getRow(1).getCell(2).value).toBe("Name");
      expect(col.header).toBe("Name");
    });

    it("a string array writes rows 1..N", () => {
      const col = host.getColumn(1);
      col.header = ["Group", "Subtotal"];
      expect(host.getRow(1).getCell(1).value).toBe("Group");
      expect(host.getRow(2).getCell(1).value).toBe("Subtotal");
    });
  });

  describe("width / hidden / outlineLevel / collapsed", () => {
    it("isCustomWidth is false at the default width, true otherwise", () => {
      const col = host.getColumn(1);
      expect(col.isCustomWidth).toBe(false);
      col.width = DEFAULT_COLUMN_WIDTH;
      expect(col.isCustomWidth).toBe(false);
      col.width = 20;
      expect(col.isCustomWidth).toBe(true);
    });

    it("collapsed derives from outlineLevel vs. the host's outlineLevelCol threshold", () => {
      const col = host.getColumn(1);
      expect(col.collapsed).toBe(false);
      col.outlineLevel = 1;
      expect(col.collapsed).toBe(true);
    });
  });

  describe("isDefault", () => {
    it("true for a bare column", () => {
      expect(host.getColumn(1).isDefault).toBe(true);
    });

    it.each([
      [
        "width",
        () => {
          host.getColumn(1).width = 30;
        },
      ],
      [
        "hidden",
        () => {
          host.getColumn(1).hidden = true;
        },
      ],
      [
        "outlineLevel",
        () => {
          host.getColumn(1).outlineLevel = 1;
        },
      ],
      [
        "key",
        () => {
          host.getColumn(1).key = "id";
        },
      ],
      [
        "header",
        () => {
          host.getColumn(1).header = "Id";
        },
      ],
      [
        "style",
        () => {
          host.getColumn(1).font = { bold: true };
        },
      ],
    ] as const)("false once %s is customized", (_label, mutate) => {
      mutate();
      expect(host.getColumn(1).isDefault).toBe(false);
    });
  });

  describe("values", () => {
    it("get returns a sparse array indexed by row number", () => {
      const col = host.getColumn(1);
      host.getRow(1).getCell(1).value = "a";
      host.getRow(3).getCell(1).value = "c";

      const values = col.values as unknown[];
      expect(values[1]).toBe("a");
      expect(values[3]).toBe("c");
      expect(values[2]).toBeUndefined();
    });

    it("set writes down the column by row number", () => {
      const col = host.getColumn(2);
      const vals: string[] = [];
      vals[1] = "x";
      vals[2] = "y";
      col.values = vals;

      expect(host.getRow(1).getCell(2).value).toBe("x");
      expect(host.getRow(2).getCell(2).value).toBe("y");
    });
  });

  describe("eachCell", () => {
    it("without includeEmpty visits only rows with a cell in this column", () => {
      const col = host.getColumn(1);
      host.getRow(1).getCell(1).value = "a";
      host.getRow(3).getCell(1).value = "c";
      host.getRow(2).getCell(2).value = "other column";

      const seen: number[] = [];
      col.eachCell((_cell, rowNumber) => seen.push(rowNumber));
      expect(seen).toEqual([1, 3]);
    });
  });

  describe("style shortcuts fan out via eachCell", () => {
    it("column.font applies to existing cells in the column only", () => {
      const col = host.getColumn(1);
      host.getRow(1).getCell(1).value = "a";
      host.getRow(1).getCell(2).value = "other column";

      col.font = { italic: true };
      expect(host.getRow(1).getCell(1).font).toEqual({ italic: true });
      expect(host.getRow(1).getCell(2).font).toBeUndefined();
      expect(col.font).toEqual({ italic: true });
    });

    it("a cell created after the column style is set inherits it", () => {
      const col = host.getColumn(1);
      col.font = { bold: true };
      const cell = host.getRow(5).getCell(1);
      expect(cell.font).toEqual({ bold: true });
    });
  });

  describe("defn", () => {
    it("round-trips key/header/width/hidden/outlineLevel/style", () => {
      const col = host.getColumn(1);
      col.key = "id";
      col.width = 15;
      col.hidden = true;
      col.outlineLevel = 1;
      col.font = { bold: true };
      col.header = "Id";

      const defn = col.defn;
      const other = host.getColumn(2);
      other.defn = defn;

      expect(other.key).toBe("id");
      expect(other.width).toBe(15);
      expect(other.hidden).toBe(true);
      expect(other.outlineLevel).toBe(1);
      expect(other.font).toEqual({ bold: true });
      expect(other.header).toBe("Id");
    });
  });

  describe("equivalentTo", () => {
    it("true for columns with the same width/hidden/outlineLevel/style", () => {
      const a = host.getColumn(1);
      const b = host.getColumn(2);
      a.width = 12;
      b.width = 12;
      a.font = { bold: true };
      b.font = { bold: true };
      expect(a.equivalentTo(b)).toBe(true);
    });

    it("false when any attribute differs", () => {
      const a = host.getColumn(1);
      const b = host.getColumn(2);
      a.width = 12;
      b.width = 20;
      expect(a.equivalentTo(b)).toBe(false);
    });

    it("ignores key/header (not OOXML-relevant)", () => {
      const a = host.getColumn(1);
      const b = host.getColumn(2);
      a.key = "id";
      b.key = "other";
      a.header = "Id";
      b.header = "Other";
      expect(a.equivalentTo(b)).toBe(true);
    });
  });

  describe("toModel / fromModel", () => {
    it("compresses adjacent equivalent columns into one range", () => {
      const cols = [1, 2, 3].map((n) => {
        const col = host.getColumn(n);
        col.width = 15;
        return col;
      });

      const model = Column.toModel(cols);
      expect(model).toEqual([{ min: 1, max: 3, width: 15 }]);
    });

    it("splits into separate ranges where columns differ", () => {
      const a = host.getColumn(1);
      a.width = 10;
      const b = host.getColumn(2);
      b.width = 20;
      const c = host.getColumn(3);
      c.width = 20;

      const model = Column.toModel([a, b, c]);
      expect(model).toEqual([
        { min: 1, max: 1, width: 10 },
        { min: 2, max: 3, width: 20 },
      ]);
    });

    it("splits a run when a column number is non-contiguous, even if equivalent", () => {
      const a = host.getColumn(1);
      a.width = 10;
      const c = host.getColumn(3);
      c.width = 10;

      const model = Column.toModel([a, c]);
      expect(model).toEqual([
        { min: 1, max: 1, width: 10 },
        { min: 3, max: 3, width: 10 },
      ]);
    });

    it("returns undefined for an empty column array", () => {
      expect(Column.toModel([])).toBeUndefined();
    });

    it("fromModel re-expands ranges back into individual columns", () => {
      const ranges: ColumnRange[] = [{ min: 2, max: 4, width: 15, hidden: true }];
      const columns = Column.fromModel(host, ranges);

      expect(columns).not.toBeNull();
      expect(columns?.map((c) => c.number)).toEqual([2, 3, 4]);
      for (const col of columns ?? []) {
        expect(col.width).toBe(15);
        expect(col.hidden).toBe(true);
      }
    });

    it("returns null for an empty range array", () => {
      expect(Column.fromModel(host, [])).toBeNull();
    });

    it("round-trips compress -> expand", () => {
      const original = [1, 2, 3, 5].map((n) => {
        const col = host.getColumn(n);
        col.width = 12;
        return col;
      });

      const model = Column.toModel(original) ?? [];
      const expanded = Column.fromModel(host, model) ?? [];

      expect(expanded.map((c) => c.number)).toEqual([1, 2, 3, 5]);
      for (const col of expanded) {
        expect(col.width).toBe(12);
      }
    });
  });
});
