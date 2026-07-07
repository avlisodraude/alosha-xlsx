import { beforeEach, describe, expect, it } from "vitest";
import type { DataValidation } from "../../src/model/data-validation.js";
import { DataValidations } from "../../src/model/data-validation.js";
import { Workbook } from "../../src/model/workbook.js";
import type { Worksheet } from "../../src/model/worksheet.js";

const listDv: DataValidation = {
  type: "list",
  allowBlank: true,
  formulae: ['"Yes,No,Maybe"'],
};

describe("DataValidations manager", () => {
  let dvs: DataValidations;

  beforeEach(() => {
    dvs = new DataValidations();
  });

  it("add / find by exact sqref key", () => {
    dvs.add("A1", listDv);
    expect(dvs.find("A1")).toEqual(listDv);
    expect(dvs.size).toBe(1);
  });

  it("find resolves an address inside a range sqref", () => {
    dvs.add("B2:B10", listDv);
    expect(dvs.find("B5")).toEqual(listDv);
    expect(dvs.find("B2")).toEqual(listDv);
    expect(dvs.find("B10")).toEqual(listDv);
    expect(dvs.find("B11")).toBeUndefined();
    expect(dvs.find("C5")).toBeUndefined();
  });

  it("find resolves an address inside a multi-range (space-separated) sqref", () => {
    dvs.add("A1:A3 C1:C3", listDv);
    expect(dvs.find("A2")).toEqual(listDv);
    expect(dvs.find("C3")).toEqual(listDv);
    expect(dvs.find("B2")).toBeUndefined();
  });

  it("find tolerates dollar-anchored ranges", () => {
    dvs.add("$E$1:$E$5", listDv);
    expect(dvs.find("E3")).toEqual(listDv);
    expect(dvs.find("E6")).toBeUndefined();
  });

  it("remove drops the entry", () => {
    dvs.add("A1", listDv);
    dvs.remove("A1");
    expect(dvs.find("A1")).toBeUndefined();
    expect(dvs.size).toBe(0);
  });

  it("add overwrites an existing sqref", () => {
    dvs.add("A1", listDv);
    const custom: DataValidation = { type: "custom", formulae: ["ISNUMBER(A1)"] };
    dvs.add("A1", custom);
    expect(dvs.find("A1")).toEqual(custom);
    expect(dvs.size).toBe(1);
  });

  it("model getter emits { sqref, ...dv } entries", () => {
    dvs.add("A1:A10", listDv);
    expect(dvs.model).toEqual([{ sqref: "A1:A10", ...listDv }]);
  });

  it("model setter clears and restores from the serialized form", () => {
    dvs.add("Z1", listDv);
    dvs.model = [
      { sqref: "A1", type: "list", formulae: ['"a,b"'] },
      { sqref: "B1:B2", type: "whole", operator: "between", formulae: [1, 10] },
    ];
    expect(dvs.size).toBe(2);
    expect(dvs.find("Z1")).toBeUndefined();
    expect(dvs.find("A1")?.type).toBe("list");
    expect(dvs.find("B2")?.operator).toBe("between");
  });
});

describe("cell.dataValidation accessor", () => {
  let ws: Worksheet;

  beforeEach(() => {
    ws = new Workbook().addWorksheet("Sheet1");
  });

  it("set routes through the worksheet manager keyed by the cell address", () => {
    ws.getCell("A1").dataValidation = listDv;
    expect(ws.dataValidations.find("A1")).toEqual(listDv);
    expect(ws.dataValidations.model).toEqual([{ sqref: "A1", ...listDv }]);
  });

  it("get reflects a validation added over a covering range", () => {
    ws.dataValidations.add("A1:A10", listDv);
    expect(ws.getCell("A5").dataValidation).toEqual(listDv);
    expect(ws.getCell("B5").dataValidation).toBeUndefined();
  });

  it("round-trips through find/add", () => {
    const cell = ws.getCell("C3");
    cell.dataValidation = { type: "custom", formulae: ["ISNUMBER(C3)"] };
    expect(cell.dataValidation).toEqual({ type: "custom", formulae: ["ISNUMBER(C3)"] });
  });
});

describe("WorksheetModel.dataValidations round-trip", () => {
  let wb: Workbook;
  let ws: Worksheet;

  beforeEach(() => {
    wb = new Workbook();
    ws = wb.addWorksheet("Sheet1");
  });

  it("omits the key when there are no validations", () => {
    expect(ws.model.dataValidations).toBeUndefined();
  });

  it("serializes validations into the model", () => {
    ws.dataValidations.add("A1:A10", listDv);
    ws.dataValidations.add("B1", { type: "decimal", operator: "between", formulae: [0, 1] });
    expect(ws.model.dataValidations).toEqual([
      { sqref: "A1:A10", ...listDv },
      { sqref: "B1", type: "decimal", operator: "between", formulae: [0, 1] },
    ]);
  });

  it("restores validations from an assigned model", () => {
    const other = wb.addWorksheet("Sheet2");
    other.dataValidations.add("D4:D8", listDv);
    const snapshot = other.model;

    const target = wb.addWorksheet("Sheet3");
    target.model = snapshot;
    expect(target.dataValidations.find("D6")).toEqual(listDv);
  });

  it("clears prior validations when a model without any is assigned", () => {
    ws.dataValidations.add("A1", listDv);
    const clean = wb.addWorksheet("Blank").model;
    ws.model = clean;
    expect(ws.dataValidations.size).toBe(0);
  });
});
