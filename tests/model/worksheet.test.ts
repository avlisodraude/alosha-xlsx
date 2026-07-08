import { beforeEach, describe, expect, it } from "vitest";
import {
  DuplicateWorksheetNameError,
  InvalidWorksheetNameError,
  MergeConflictError,
  ReservedWorksheetNameError,
  WorksheetNameTooLongError,
} from "../../src/errors.js";
import { CellType } from "../../src/model/enums.js";
import { Workbook } from "../../src/model/workbook.js";
import type { Worksheet } from "../../src/model/worksheet.js";

describe("Worksheet", () => {
  let wb: Workbook;
  let ws: Worksheet;

  beforeEach(() => {
    wb = new Workbook();
    ws = wb.addWorksheet("Sheet1");
  });

  describe("rows & cells created on demand", () => {
    it("getRow / getCell create on demand and are stable", () => {
      const row = ws.getRow(3);
      expect(row.number).toBe(3);
      expect(ws.getRow(3)).toBe(row);

      const cell = ws.getCell("B3");
      expect(cell.row).toBe(3);
      expect(cell.col).toBe(2);
      expect(ws.getCell(3, 2)).toBe(cell);
    });

    it("findCell / findRow do not create", () => {
      expect(ws.findRow(5)).toBeUndefined();
      expect(ws.findCell("A5")).toBeUndefined();
      ws.getCell("A5").value = 1;
      expect(ws.findRow(5)).toBeDefined();
      expect(ws.findCell("A5")?.value).toBe(1);
    });

    it("addRow appends at the next row", () => {
      ws.getCell("A1").value = "header";
      const r2 = ws.addRow(["a", "b", "c"]);
      expect(r2.number).toBe(2);
      expect(ws.getCell("A2").value).toBe("a");
      expect(ws.getCell("C2").value).toBe("c");

      const r3 = ws.addRow(["x"]);
      expect(r3.number).toBe(3);
    });

    it("addRows appends several rows", () => {
      const rows = ws.addRows([["a"], ["b"], ["c"]]);
      expect(rows.map((r) => r.number)).toEqual([1, 2, 3]);
      expect(ws.getCell("A3").value).toBe("c");
    });

    it("getSheetValues returns a sparse, 1-based array", () => {
      ws.getRow(1).values = ["a", "b"];
      ws.getRow(3).values = ["c"];
      const values = ws.getSheetValues();
      expect(values[0]).toBeUndefined();
      expect((values[1] as unknown[])[1]).toBe("a");
      expect((values[3] as unknown[])[1]).toBe("c");
    });
  });

  describe("eachRow", () => {
    it("visits only non-empty rows by default", () => {
      ws.getRow(1).getCell(1).value = "a";
      ws.getRow(4).getCell(1).value = "d";
      const seen: number[] = [];
      ws.eachRow((_row, n) => seen.push(n));
      expect(seen).toEqual([1, 4]);
    });

    it("includeEmpty visits every row up to rowCount", () => {
      ws.getRow(1).getCell(1).value = "a";
      ws.getRow(3).getCell(1).value = "c";
      const seen: number[] = [];
      ws.eachRow({ includeEmpty: true }, (_row, n) => seen.push(n));
      expect(seen).toEqual([1, 2, 3]);
    });
  });

  describe("counts & dimensions", () => {
    it("rowCount / actualRowCount / lastRow", () => {
      ws.getRow(1).getCell(1).value = "a";
      ws.getRow(4); // created but empty
      ws.getRow(3).getCell(1).value = "c";
      expect(ws.rowCount).toBe(4);
      expect(ws.actualRowCount).toBe(2);
      expect(ws.lastRow?.number).toBe(4);
    });

    it("columnCount / actualColumnCount / lastColumn", () => {
      ws.getCell("A1").value = "a";
      ws.getCell("D1").value = "d";
      expect(ws.columnCount).toBe(4);
      expect(ws.actualColumnCount).toBe(2);
      expect(ws.lastColumn?.number).toBe(4);
    });

    it("dimensions is the bounding box of used cells", () => {
      ws.getCell("B2").value = "x";
      ws.getCell("D5").value = "y";
      expect(ws.dimensions.range).toBe("B2:D5");
    });
  });

  describe("merges", () => {
    it("mergeCells links slaves to the master; hasMerges reflects it", () => {
      expect(ws.hasMerges).toBe(false);
      ws.getCell("A1").value = "title";
      ws.mergeCells("A1:B2");
      expect(ws.hasMerges).toBe(true);

      const slave = ws.getCell("B2");
      expect(slave.type).toBe(CellType.Merge);
      expect(slave.isMerged).toBe(true);
      expect(slave.master).toBe(ws.getCell("A1"));
      // writing a slave writes the master
      slave.value = "changed";
      expect(ws.getCell("A1").value).toBe("changed");
    });

    it("accepts the (top,left,bottom,right) and two-address forms", () => {
      ws.mergeCells(1, 1, 2, 2);
      expect(ws.getCell("B2").isMerged).toBe(true);

      ws.mergeCells("C1", "D1");
      expect(ws.getCell("D1").isMerged).toBe(true);
    });

    it("throws MergeConflictError on an overlapping merge and leaves state untouched", () => {
      ws.mergeCells("A1:B2");
      expect(() => ws.mergeCells("B2:C3")).toThrow(MergeConflictError);
      // the failed merge must not have partially applied
      expect(ws.getCell("C3").isMerged).toBe(false);
    });

    it("unMergeCells unmerges the whole block and preserves the master value", () => {
      ws.getCell("A1").value = "keep";
      ws.mergeCells("A1:B2");
      ws.unMergeCells("B2");
      expect(ws.hasMerges).toBe(false);
      expect(ws.getCell("B2").isMerged).toBe(false);
      expect(ws.getCell("A1").value).toBe("keep");
    });

    it("authoring-time mergeCells propagates the master's style onto its slaves", () => {
      ws.getCell("A1").font = { bold: true };
      ws.getCell("B2").border = { bottom: { style: "thick" } };
      ws.mergeCells("A1:B2");
      // The slave adopts the master's style wholesale, dropping its own border (ExcelJS behavior).
      expect(ws.getCell("B2").style).toEqual(ws.getCell("A1").style);
      expect(ws.getCell("B2").border).toBeUndefined();
    });

    it("mergeCellsPreservingStyles keeps each slave's own style (read/model re-hydration path)", () => {
      ws.getCell("A1").font = { bold: true };
      ws.getCell("B2").border = { bottom: { style: "thick" } };
      ws.mergeCellsPreservingStyles("A1:B2");
      expect(ws.getCell("B2").isMerged).toBe(true);
      expect(ws.getCell("B2").border).toEqual({ bottom: { style: "thick" } });
    });
  });

  describe("name validation", () => {
    it("rejects an empty name", () => {
      expect(() => {
        ws.name = "";
      }).toThrow(InvalidWorksheetNameError);
    });

    it("rejects illegal characters", () => {
      for (const bad of ["a/b", "a\\b", "a?b", "a*b", "a:b", "a[b", "a]b"]) {
        expect(() => {
          ws.name = bad;
        }).toThrow(InvalidWorksheetNameError);
      }
    });

    it("rejects a leading or trailing single quote", () => {
      expect(() => {
        ws.name = "'quoted";
      }).toThrow(InvalidWorksheetNameError);
      expect(() => {
        ws.name = "quoted'";
      }).toThrow(InvalidWorksheetNameError);
    });

    it("rejects the reserved name History (case-insensitive)", () => {
      expect(() => {
        ws.name = "History";
      }).toThrow(ReservedWorksheetNameError);
      expect(() => {
        ws.name = "history";
      }).toThrow(ReservedWorksheetNameError);
    });

    it("rejects a duplicate name (case-insensitive)", () => {
      wb.addWorksheet("Other");
      expect(() => {
        ws.name = "other";
      }).toThrow(DuplicateWorksheetNameError);
    });

    it("allows renaming a sheet to its own name", () => {
      expect(() => {
        ws.name = "Sheet1";
      }).not.toThrow();
    });

    it("rejects a name longer than 31 characters", () => {
      expect(() => {
        ws.name = "x".repeat(32);
      }).toThrow(WorksheetNameTooLongError);
    });

    it("accepts a valid rename", () => {
      ws.name = "Renamed";
      expect(ws.name).toBe("Renamed");
    });
  });

  describe("column keys", () => {
    it("registers and resolves keys", () => {
      const col = ws.getColumn(3);
      ws.setColumnKey("total", col);
      expect(ws.getColumnKey("total")).toBe(col);
      // addRow with an object uses column keys
      ws.addRow({ total: 42 });
      expect(ws.getCell("C1").value).toBe(42);

      ws.deleteColumnKey("total");
      expect(ws.getColumnKey("total")).toBeUndefined();
    });

    it("columns setter assigns definitions and keys", () => {
      ws.columns = [
        { header: "Id", key: "id", width: 10 },
        { header: "Name", key: "name", width: 20 },
      ];
      expect(ws.getColumn(1).width).toBe(10);
      expect(ws.getColumnKey("name")?.number).toBe(2);
      expect(ws.getCell("A1").value).toBe("Id");
    });
  });

  describe("spliceRows", () => {
    it("inserts rows and shifts the ones below down", () => {
      ws.getRow(1).values = ["a"];
      ws.getRow(2).values = ["b"];
      ws.spliceRows(2, 0, ["inserted"]);
      expect(ws.getCell("A1").value).toBe("a");
      expect(ws.getCell("A2").value).toBe("inserted");
      expect(ws.getCell("A3").value).toBe("b");
    });

    it("deletes rows and shifts the ones below up, carrying styles", () => {
      ws.getRow(1).values = ["a"];
      ws.getRow(2).values = ["b"];
      ws.getRow(3).values = ["c"];
      ws.getCell("A3").font = { bold: true };
      ws.spliceRows(2, 1);
      expect(ws.getCell("A1").value).toBe("a");
      expect(ws.getCell("A2").value).toBe("c");
      expect(ws.getCell("A2").font).toEqual({ bold: true });
      expect(ws.findCell("A3")).toBeUndefined();
    });

    it("re-anchors a merge that moves when rows are inserted above it", () => {
      ws.getCell("B3").value = "m";
      ws.mergeCells("B3:C4");
      // insert two rows at the top -> the merge shifts down by 2 to B5:C6
      ws.spliceRows(1, 0, ["x"], ["y"]);

      expect(ws.getCell("B5").isMerged).toBe(false);
      expect(ws.getCell("B5").master).toBe(ws.getCell("B5"));
      expect(ws.getCell("C6").isMerged).toBe(true);
      expect(ws.getCell("C6").master).toBe(ws.getCell("B5"));
      expect(ws.hasMerges).toBe(true);
    });

    it("re-anchors a merge that moves up when rows are deleted above it", () => {
      ws.getCell("A5").value = "m";
      ws.mergeCells("A5:B6");
      ws.spliceRows(1, 2); // delete rows 1-2 -> merge shifts up to A3:B4
      expect(ws.getCell("B4").isMerged).toBe(true);
      expect(ws.getCell("B4").master).toBe(ws.getCell("A3"));
    });

    it("shifts a comment, data validation and conditional format below an inserted row", () => {
      ws.comments.add("A5", { texts: "note" });
      ws.dataValidations.add("A5:A7", { type: "list", formulae: ['"a,b"'] });
      ws.addConditionalFormatting({
        ref: "A5:A7",
        rules: [{ type: "expression", formulae: ["A5>0"] }],
      });

      ws.spliceRows(2, 0, ["x"]); // insert one row above -> everything shifts down by 1

      expect(ws.comments.find("A6")).toEqual({ texts: "note" });
      expect(ws.comments.find("A5")).toBeUndefined();
      expect(ws.dataValidations.model[0]?.sqref).toBe("A6:A8");
      expect(ws.conditionalFormattings[0]?.ref).toBe("A6:A8");
    });

    it("keeps annotations entirely above the removed region untouched", () => {
      ws.comments.add("A1", { texts: "top" });
      ws.dataValidations.add("A1", { type: "whole", formulae: [1] });
      ws.spliceRows(5, 1); // delete row 5, far below the annotations
      expect(ws.comments.find("A1")).toEqual({ texts: "top" });
      expect(ws.dataValidations.model[0]?.sqref).toBe("A1");
    });

    it("drops an annotation straddling the removed rows, mirroring merges", () => {
      ws.dataValidations.add("A4:A6", { type: "whole", formulae: [1] });
      ws.comments.add("A5", { texts: "gone" });
      ws.spliceRows(5, 1); // delete row 5 -> the single-cell comment vanishes, the range is dropped
      expect(ws.comments.find("A5")).toBeUndefined();
      expect(ws.dataValidations.size).toBe(0);
    });

    it("preserves independent parts of a space-separated sqref, dropping only the straddling one", () => {
      ws.dataValidations.add("A1 A5 A9", { type: "whole", formulae: [1] });
      ws.spliceRows(5, 1); // delete row 5 -> A1 kept, A5 dropped, A9 shifts to A8
      expect(ws.dataValidations.model[0]?.sqref).toBe("A1 A8");
    });

    it("leaves images and autoFilter at their original addresses (ExcelJS parity)", () => {
      const imageId = wb.addImage({ extension: "png", buffer: new Uint8Array([1, 2, 3]) });
      ws.addImage(imageId, "B5:D8"); // stored 0-based: tl col1/row4, br col3/row7
      ws.autoFilter = "A5:C5";
      ws.spliceRows(1, 0, ["x"]); // insert a row at the top
      expect(ws.images[0]?.range).toEqual({ tl: { col: 1, row: 4 }, br: { col: 3, row: 7 } });
      expect(ws.autoFilter).toBe("A5:C5");
    });
  });

  describe("spliceColumns", () => {
    it("uses 1-based insert arrays (colVals[n] fills row n, index 0 unused)", () => {
      ws.getCell("A1").value = "a1";
      ws.getCell("A2").value = "a2";
      // Insert one column at A: index 1 -> row 1, index 2 -> row 2, index 0 (null) ignored.
      ws.spliceColumns(1, 0, [null, "top", "bottom"]);
      expect(ws.getCell("A1").value).toBe("top");
      expect(ws.getCell("A2").value).toBe("bottom");
      expect(ws.getCell("B1").value).toBe("a1");
      expect(ws.getCell("B2").value).toBe("a2");
    });

    it("shifts a comment, data validation and conditional format right of an inserted column", () => {
      ws.comments.add("C1", { texts: "note" });
      ws.dataValidations.add("C1:E1", { type: "list", formulae: ['"a,b"'] });
      ws.addConditionalFormatting({
        ref: "C1:E1",
        rules: [{ type: "expression", formulae: ["C1>0"] }],
      });

      ws.spliceColumns(1, 0, [null]); // insert one column at A -> everything shifts right by 1

      expect(ws.comments.find("D1")).toEqual({ texts: "note" });
      expect(ws.dataValidations.model[0]?.sqref).toBe("D1:F1");
      expect(ws.conditionalFormattings[0]?.ref).toBe("D1:F1");
    });

    it("leaves images and autoFilter in place on a column splice (ExcelJS parity)", () => {
      const imageId = wb.addImage({ extension: "png", buffer: new Uint8Array([1, 2, 3]) });
      ws.addImage(imageId, "C1:E4"); // stored 0-based: tl col2/row0, br col4/row3
      ws.autoFilter = "C1:E1";
      ws.spliceColumns(1, 0, [null]);
      expect(ws.images[0]?.range).toEqual({ tl: { col: 2, row: 0 }, br: { col: 4, row: 3 } });
      expect(ws.autoFilter).toBe("C1:E1");
    });
  });

  describe("insertRow / duplicateRow", () => {
    it("insertRow inserts at a position and returns the new row", () => {
      ws.getRow(1).values = ["a"];
      ws.getRow(2).values = ["b"];
      const row = ws.insertRow(2, ["new"]);
      expect(row.number).toBe(2);
      expect(ws.getCell("A2").value).toBe("new");
      expect(ws.getCell("A3").value).toBe("b");
    });

    it("duplicateRow copies values into following rows", () => {
      ws.getRow(1).values = ["a", "b"];
      ws.getCell("A1").font = { italic: true };
      ws.duplicateRow(1, 2);
      expect(ws.getCell("A2").value).toBe("a");
      expect(ws.getCell("B3").value).toBe("b");
      expect(ws.getCell("A2").font).toEqual({ italic: true });
    });
  });

  describe("model round-trip", () => {
    it("round-trips rows, columns and merges through get/set model", () => {
      ws.columns = [{ key: "id", width: 12 }];
      ws.getCell("A1").value = "keep";
      ws.getRow(1).values = ["keep", "next"];
      ws.getCell("C3").value = 99;
      ws.mergeCells("A1:B1");

      const snapshot = ws.model;

      const other = wb.addWorksheet("Clone");
      other.model = snapshot;

      expect(other.model).toEqual(snapshot);
      expect(other.getCell("B1").isMerged).toBe(true);
      expect(other.getCell("C3").value).toBe(99);
    });
  });
});
