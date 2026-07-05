import { beforeEach, describe, expect, it } from "vitest";
import { DuplicateWorksheetNameError } from "../../src/errors.js";
import { Workbook } from "../../src/model/workbook.js";

describe("Workbook", () => {
  let wb: Workbook;

  beforeEach(() => {
    wb = new Workbook();
  });

  describe("addWorksheet", () => {
    it("auto-names sheets and assigns incrementing ids and order", () => {
      const a = wb.addWorksheet();
      const b = wb.addWorksheet();
      expect(a.name).toBe("Sheet1");
      expect(b.name).toBe("Sheet2");
      expect(a.id).toBe(1);
      expect(b.id).toBe(2);
      expect(a.orderNo).toBe(0);
      expect(b.orderNo).toBe(1);
    });

    it("avoids collisions when auto-naming around an explicit name", () => {
      wb.addWorksheet("Sheet1");
      const auto = wb.addWorksheet();
      expect(auto.name).toBe("Sheet2");
    });

    it("rejects a duplicate explicit name (case-insensitive)", () => {
      wb.addWorksheet("Data");
      expect(() => wb.addWorksheet("data")).toThrow(DuplicateWorksheetNameError);
    });

    it("applies options (state / properties)", () => {
      const ws = wb.addWorksheet("Hidden", {
        state: "hidden",
        properties: { defaultRowHeight: 20 },
      });
      expect(ws.state).toBe("hidden");
      expect(ws.properties.defaultRowHeight).toBe(20);
    });
  });

  describe("getWorksheet", () => {
    it("looks up by id and by name, and defaults to the first sheet", () => {
      const a = wb.addWorksheet("A");
      const b = wb.addWorksheet("B");
      expect(wb.getWorksheet(a.id)).toBe(a);
      expect(wb.getWorksheet("B")).toBe(b);
      expect(wb.getWorksheet()).toBe(a);
      expect(wb.getWorksheet("missing")).toBeUndefined();
      expect(wb.getWorksheet(999)).toBeUndefined();
    });
  });

  describe("worksheets ordering", () => {
    it("stays ordered by insertion", () => {
      wb.addWorksheet("A");
      wb.addWorksheet("B");
      wb.addWorksheet("C");
      expect(wb.worksheets.map((w) => w.name)).toEqual(["A", "B", "C"]);
    });

    it("re-sequences orderNo after a removal", () => {
      const a = wb.addWorksheet("A");
      wb.addWorksheet("B");
      const c = wb.addWorksheet("C");
      wb.removeWorksheet("B");

      expect(wb.worksheets.map((w) => w.name)).toEqual(["A", "C"]);
      expect(a.orderNo).toBe(0);
      expect(c.orderNo).toBe(1);
      expect(wb.getWorksheet("B")).toBeUndefined();
    });

    it("worksheets returns a copy that cannot mutate the backing order", () => {
      wb.addWorksheet("A");
      const list = wb.worksheets;
      list.pop();
      expect(wb.worksheets.map((w) => w.name)).toEqual(["A"]);
    });
  });

  describe("removeWorksheet", () => {
    it("removes by id and is a no-op for an unknown handle", () => {
      const a = wb.addWorksheet("A");
      wb.addWorksheet("B");
      wb.removeWorksheet(a.id);
      expect(wb.worksheets.map((w) => w.name)).toEqual(["B"]);
      expect(() => wb.removeWorksheet("nope")).not.toThrow();
    });
  });

  describe("eachSheet", () => {
    it("iterates sheets with their ids", () => {
      const a = wb.addWorksheet("A");
      const b = wb.addWorksheet("B");
      const seen: Array<[string, number]> = [];
      wb.eachSheet((ws, id) => seen.push([ws.name, id]));
      expect(seen).toEqual([
        ["A", a.id],
        ["B", b.id],
      ]);
    });
  });

  describe("metadata", () => {
    it("holds document properties", () => {
      wb.creator = "alosha";
      wb.title = "Report";
      wb.properties.date1904 = true;
      expect(wb.creator).toBe("alosha");
      expect(wb.title).toBe("Report");
      expect(wb.properties.date1904).toBe(true);
    });
  });

  describe("model round-trip", () => {
    it("deep-equals through get model -> set model", () => {
      wb.creator = "alosha";
      wb.title = "Report";
      wb.properties.date1904 = true;

      const s1 = wb.addWorksheet("Data");
      s1.columns = [{ key: "id", width: 12 }];
      s1.getRow(1).values = ["a", "b"];
      s1.getCell("C3").value = 42;
      s1.mergeCells("A1:B1");

      const s2 = wb.addWorksheet("Summary", { state: "hidden" });
      s2.getCell("A1").value = "total";

      const snapshot = wb.model;

      const clone = new Workbook();
      clone.model = snapshot;

      expect(clone.model).toEqual(snapshot);
      expect(clone.worksheets.map((w) => w.name)).toEqual(["Data", "Summary"]);
      expect(clone.getWorksheet("Data")?.getCell("C3").value).toBe(42);
      expect(clone.getWorksheet("Data")?.getCell("B1").isMerged).toBe(true);
      expect(clone.getWorksheet("Summary")?.state).toBe("hidden");
    });
  });
});
