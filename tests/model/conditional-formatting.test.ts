import { beforeEach, describe, expect, it } from "vitest";
import { InvalidRangeError } from "../../src/errors.js";
import type {
  ConditionalFormatting,
  ConditionalRule,
} from "../../src/model/conditional-formatting.js";
import { Workbook } from "../../src/model/workbook.js";
import type { Worksheet } from "../../src/model/worksheet.js";

const cellIs: ConditionalFormatting = {
  ref: "A1:A10",
  rules: [
    {
      type: "cellIs",
      operator: "greaterThan",
      formulae: [5],
      style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFF8696B" } } },
    },
  ],
};

const colorScale: ConditionalFormatting = {
  ref: "C1:C20",
  rules: [
    {
      type: "colorScale",
      cfvo: [{ type: "min" }, { type: "max" }],
      color: ["FFF8696B", "FF63BE7B"],
    },
  ],
};

describe("Worksheet conditional-formatting manager", () => {
  let ws: Worksheet;

  beforeEach(() => {
    ws = new Workbook().addWorksheet("Sheet1");
  });

  it("starts empty", () => {
    expect(ws.conditionalFormattings).toEqual([]);
  });

  it("add appends and returns the block; getter reflects insertion order", () => {
    expect(ws.addConditionalFormatting(cellIs)).toBe(cellIs);
    ws.addConditionalFormatting(colorScale);
    expect(ws.conditionalFormattings).toEqual([cellIs, colorScale]);
  });

  it("remove drops every block with the matching ref", () => {
    ws.addConditionalFormatting(cellIs);
    ws.addConditionalFormatting(colorScale);
    ws.removeConditionalFormatting("A1:A10");
    expect(ws.conditionalFormattings).toEqual([colorScale]);
  });

  it("remove of an unknown ref is a no-op", () => {
    ws.addConditionalFormatting(cellIs);
    ws.removeConditionalFormatting("Z9:Z9");
    expect(ws.conditionalFormattings).toEqual([cellIs]);
  });

  it("accepts a space-separated multi-range ref", () => {
    const cf: ConditionalFormatting = { ref: "A1:A3 C1:C3", rules: [] };
    expect(() => ws.addConditionalFormatting(cf)).not.toThrow();
    expect(ws.conditionalFormattings).toEqual([cf]);
  });

  it("throws a typed InvalidRangeError on a malformed ref", () => {
    expect(() => ws.addConditionalFormatting({ ref: "A1:", rules: [] })).toThrow(InvalidRangeError);
    expect(() => ws.addConditionalFormatting({ ref: "@@@", rules: [] })).toThrow(InvalidRangeError);
    // nothing was stored on the failed add
    expect(ws.conditionalFormattings).toEqual([]);
  });
});

describe("WorksheetModel.conditionalFormattings round-trip", () => {
  let wb: Workbook;
  let ws: Worksheet;

  beforeEach(() => {
    wb = new Workbook();
    ws = wb.addWorksheet("Sheet1");
  });

  it("omits the key when there are no blocks", () => {
    expect(ws.model.conditionalFormattings).toBeUndefined();
  });

  it("serializes blocks into the model", () => {
    ws.addConditionalFormatting(cellIs);
    ws.addConditionalFormatting(colorScale);
    expect(ws.model.conditionalFormattings).toEqual([cellIs, colorScale]);
  });

  it("model getter emits a deep copy (mutating the snapshot leaves the sheet intact)", () => {
    ws.addConditionalFormatting(cellIs);
    const snap = ws.model.conditionalFormattings;
    expect(snap).toBeDefined();
    snap?.[0]?.rules.push({ type: "aboveAverage" });
    expect(ws.conditionalFormattings[0]?.rules).toHaveLength(1);
  });

  it("restores blocks from an assigned model", () => {
    const other = wb.addWorksheet("Sheet2");
    other.addConditionalFormatting(cellIs);
    other.addConditionalFormatting(colorScale);
    const snapshot = other.model;

    const target = wb.addWorksheet("Sheet3");
    target.model = snapshot;
    expect(target.conditionalFormattings).toEqual([cellIs, colorScale]);
  });

  it("clears prior blocks when a model without any is assigned", () => {
    ws.addConditionalFormatting(cellIs);
    const clean = wb.addWorksheet("Blank").model;
    ws.model = clean;
    expect(ws.conditionalFormattings).toEqual([]);
  });
});

describe("ConditionalRule union narrows on type", () => {
  it("discriminates each in-scope rule shape", () => {
    const rules: ConditionalRule[] = [
      { type: "cellIs", operator: "between", formulae: [1, 10] },
      { type: "expression", formulae: ["$A1>0"] },
      { type: "containsText", text: "foo" },
      { type: "top10", rank: 10, percent: true, bottom: true },
      { type: "aboveAverage" },
      { type: "duplicateValues" },
      {
        type: "colorScale",
        cfvo: [{ type: "min" }, { type: "max" }],
        color: ["FFF8696B", "FF63BE7B"],
      },
      { type: "dataBar", cfvo: [{ type: "min" }, { type: "max" }], color: "FF638EC6" },
      { type: "iconSet", iconSet: "3TrafficLights1", cfvo: [{ type: "percent", value: 0 }] },
    ];

    for (const rule of rules) {
      switch (rule.type) {
        case "cellIs":
          expect(rule.operator).toBe("between");
          expect(rule.formulae).toEqual([1, 10]);
          break;
        case "expression":
          expect(rule.formulae[0]).toBe("$A1>0");
          break;
        case "containsText":
        case "notContainsText":
        case "beginsWith":
        case "endsWith":
          expect(rule.text).toBe("foo");
          break;
        case "top10":
          expect(rule.rank).toBe(10);
          expect(rule.percent).toBe(true);
          expect(rule.bottom).toBe(true);
          break;
        case "aboveAverage":
        case "belowAverage":
          expect(rule.type).toBe("aboveAverage");
          break;
        case "duplicateValues":
        case "uniqueValues":
          expect(rule.type).toBe("duplicateValues");
          break;
        case "colorScale":
          expect(rule.color).toHaveLength(2);
          expect(rule.cfvo).toHaveLength(2);
          break;
        case "dataBar":
          expect(rule.color).toBe("FF638EC6");
          break;
        case "iconSet":
          expect(rule.iconSet).toBe("3TrafficLights1");
          expect(rule.cfvo[0]?.value).toBe(0);
          break;
      }
    }
  });
});
