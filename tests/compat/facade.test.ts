import ExcelJS, { UnsupportedFeatureError, ValueType, Workbook } from "@alosha/xlsx/compat";
import { describe, expect, it } from "vitest";
import { readWorkbookBuffer } from "../../src/index.js";

// The facade contract from COMPAT-SPEC.md §5: an ExcelJS-faithful surface over the core model. These
// exercise the three documented divergences (cell.type numbers, whole-object style precedence,
// row/column/value parity) and the deferred surface (§6), without needing the real `exceljs` dep —
// the cross-check against exceljs itself is the parity corpus (session 4).

describe("compat facade — cell.type / effectiveType report ExcelJS ValueType numbers (§3, D1)", () => {
  it("maps every value kind to its ExcelJS number", () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("S");

    ws.getCell("A1").value = null;
    ws.getCell("A2").value = 42;
    ws.getCell("A3").value = "hi";
    ws.getCell("A4").value = true;
    ws.getCell("A5").value = new Date("2020-01-02T00:00:00Z");
    ws.getCell("A6").value = { formula: "B1+1", result: 5 };
    ws.getCell("A7").value = { text: "site", hyperlink: "https://example.com" };
    ws.getCell("A8").value = { richText: [{ text: "rich" }] };
    ws.getCell("A9").value = { error: "#N/A" };

    expect(ws.getCell("A1").type).toBe(ValueType.Null); // 0
    expect(ws.getCell("A2").type).toBe(ValueType.Number); // 2
    expect(ws.getCell("A3").type).toBe(ValueType.String); // 3
    expect(ws.getCell("A4").type).toBe(ValueType.Boolean); // 9
    expect(ws.getCell("A5").type).toBe(ValueType.Date); // 4
    expect(ws.getCell("A6").type).toBe(ValueType.Formula); // 6
    expect(ws.getCell("A7").type).toBe(ValueType.Hyperlink); // 5
    expect(ws.getCell("A8").type).toBe(ValueType.RichText); // 8
    expect(ws.getCell("A9").type).toBe(ValueType.Error); // 10
  });

  it("reports Merge (1) for a merged slave cell", () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = "master";
    ws.mergeCells("A1:B1");
    expect(ws.getCell("B1").type).toBe(ValueType.Merge); // 1
  });

  it("effectiveType of a formula follows its cached result", () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = { formula: "1+1", result: 2 };
    expect(ws.getCell("A1").type).toBe(ValueType.Formula); // 6
    expect(ws.getCell("A1").effectiveType).toBe(ValueType.Number); // 2
  });

  it("hands proxied cells to eachRow/eachCell callbacks (type stays mapped)", () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = 7;
    const seen: number[] = [];
    ws.eachRow((row) => {
      row.eachCell((cell) => seen.push(cell.type));
    });
    expect(seen).toEqual([ValueType.Number]);
  });
});

describe("compat facade — whole-object style precedence via applyExcelJsStyle (§4)", () => {
  it("replaces the font category outright rather than field-merging an inherited font", () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("S");
    ws.getRow(1).font = { name: "Arial", size: 12 };
    const cell = ws.getCell("A1");
    expect(cell.font).toEqual({ name: "Arial", size: 12 }); // inherited

    cell.font = { bold: true };

    expect(cell.font).toEqual({ bold: true }); // name/size discarded — ExcelJS semantics
    expect(cell.font).not.toHaveProperty("name");
  });

  it("clears a category (and its style key) when set to undefined", () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("S");
    const cell = ws.getCell("A1");
    cell.font = { name: "Arial", size: 12 };
    cell.font = undefined;
    expect(cell.font).toBeUndefined();
    expect(cell.style).not.toHaveProperty("font");
  });

  it("routes numFmt through the same whole-object path", () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("S");
    const cell = ws.getCell("A1");
    cell.numFmt = "0.00%";
    expect(cell.numFmt).toBe("0.00%");
  });
});

describe("compat facade — row / column / value parity (§5)", () => {
  it("addRow(array) is 1-based; row.values is 1-based with an empty [0]", () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("S");
    const row = ws.addRow(["a", "b", "c"]);
    expect(row.getCell(1).value).toBe("a");
    expect(row.getCell(3).value).toBe("c");
    const values = row.values as unknown[];
    expect(values[0]).toBeUndefined();
    expect(values[1]).toBe("a");
    expect(values[3]).toBe("c");
  });

  it("worksheet.columns = [{header,key,width}] then addRow(object) keys by column", () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("S");
    ws.columns = [
      { header: "Name", key: "name", width: 20 },
      { header: "Age", key: "age", width: 10 },
    ];
    expect(ws.getCell("A1").value).toBe("Name"); // header written to row 1
    expect(ws.getColumn("name").width).toBe(20);

    const row = ws.addRow({ name: "Ada", age: 36 });
    expect(row.getCell("name").value).toBe("Ada");
    expect(row.getCell("age").value).toBe(36);
    expect(ws.getCell("A2").value).toBe("Ada");
    expect(ws.getCell("B2").value).toBe(36);
  });
});

describe("compat facade — images bridge straight to the core registry (IMAGES-DRAWINGS-SPEC.md §4)", () => {
  it("workbook.addImage/getImage/media and worksheet.addImage/getImages round-trip in memory", () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("S");
    const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47]);

    const workbookApi = wb as unknown as {
      addImage: (img: { extension: string; buffer: Uint8Array }) => number;
      getImage: (id: number) => { type: string; extension: string; buffer: Uint8Array } | undefined;
      media: Array<{ type: string; extension: string; buffer: Uint8Array }>;
    };
    const worksheetApi = ws as unknown as {
      addImage: (id: number, anchor: string) => void;
      getImages: () => Array<{ type: string; imageId: number; range: unknown }>;
    };

    const imageId = workbookApi.addImage({ extension: "png", buffer: png });
    expect(imageId).toBe(0);
    expect(workbookApi.getImage(imageId)).toEqual({ type: "image", extension: "png", buffer: png });
    expect(workbookApi.media).toEqual([{ type: "image", extension: "png", buffer: png }]);

    worksheetApi.addImage(imageId, "B2:D6");
    expect(worksheetApi.getImages()).toEqual([
      {
        type: "image",
        imageId: 0,
        range: { tl: { col: 1, row: 1 }, br: { col: 3, row: 5 } },
      },
    ]);
  });
});

describe("compat facade — identity & instanceof", () => {
  it("is a real Workbook and returns stable cell facades", () => {
    const wb = new ExcelJS.Workbook();
    expect(wb).toBeInstanceOf(Workbook);
    const ws = wb.addWorksheet("S");
    expect(ws.getCell("A1")).toBe(ws.getCell("A1")); // stable identity
  });

  it("round-trips through workbook.xlsx and keeps ExcelJS types after load", async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Data");
    ws.getCell("A1").value = 3.14;
    ws.getCell("B1").value = "pi";
    const buf = await wb.xlsx.writeBuffer();

    const reloaded = await new ExcelJS.Workbook().xlsx.load(buf);
    const rws = reloaded.getWorksheet("Data");
    expect(rws?.getCell("A1").type).toBe(ValueType.Number);
    expect(rws?.getCell("B1").type).toBe(ValueType.String);
    expect(rws?.getCell("A1").value).toBe(3.14);
  });
});

describe("compat facade — deferred surface throws UnsupportedFeatureError (§6)", () => {
  it("workbook-level: csv / definedNames", () => {
    const wb = new ExcelJS.Workbook();
    expect(() => (wb as unknown as { csv: { writeBuffer: () => void } }).csv.writeBuffer()).toThrow(
      UnsupportedFeatureError,
    );
    expect(() =>
      (wb as unknown as { definedNames: { add: () => void } }).definedNames.add(),
    ).toThrow(UnsupportedFeatureError);
  });

  it("workbook-level: addImage rejects base64/filename inputs and a missing buffer", () => {
    const wb = new ExcelJS.Workbook();
    const addImage = (wb as unknown as { addImage: (img: unknown) => number }).addImage;
    expect(() => addImage({ extension: "png", base64: "AA==" })).toThrow(UnsupportedFeatureError);
    expect(() => addImage({ extension: "png", filename: "a.png" })).toThrow(
      UnsupportedFeatureError,
    );
    expect(() => addImage({ extension: "png" })).toThrow(UnsupportedFeatureError);
  });

  it("worksheet-level: addTable / protect", () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("S") as unknown as {
      addTable: () => void;
      protect: () => void;
    };
    expect(() => ws.addTable()).toThrow(UnsupportedFeatureError);
    expect(() => ws.protect()).toThrow(UnsupportedFeatureError);
  });

  it("worksheet-level: addBackgroundImage stays a documented stub", () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("S") as unknown as { addBackgroundImage: () => void };
    expect(() => ws.addBackgroundImage()).toThrow(UnsupportedFeatureError);
  });

  it("namespace-level: stream.xlsx.WorkbookWriter / WorkbookReader", () => {
    expect(() => new ExcelJS.stream.xlsx.WorkbookWriter()).toThrow(UnsupportedFeatureError);
    expect(() => new ExcelJS.stream.xlsx.WorkbookReader()).toThrow(UnsupportedFeatureError);
  });
});

describe("compat facade — data validation is a real pass-through (DATA-VALIDATION-SPEC §5)", () => {
  it("cell.dataValidation get/set reaches the core accessor", () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("S");
    const cell = ws.getCell("A1") as unknown as {
      dataValidation: {
        type: string;
        allowBlank?: boolean;
        formulae?: (string | number)[];
      };
    };

    cell.dataValidation = { type: "list", allowBlank: true, formulae: ['"Yes,No"'] };
    expect(cell.dataValidation).toEqual({
      type: "list",
      allowBlank: true,
      formulae: ['"Yes,No"'],
    });
  });

  it("worksheet.dataValidations exposes the ExcelJS-shaped manager (add/find/remove/model)", () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("S") as unknown as {
      dataValidations: {
        add: (sqref: string, dv: unknown) => unknown;
        find: (address: string) => unknown;
        remove: (sqref: string) => void;
        model: unknown[];
      };
    };

    ws.dataValidations.add("B2:B10", { type: "decimal", operator: "between", formulae: [0, 1] });
    expect(ws.dataValidations.find("B5")).toEqual({
      type: "decimal",
      operator: "between",
      formulae: [0, 1],
    });
    expect(ws.dataValidations.model).toEqual([
      { sqref: "B2:B10", type: "decimal", operator: "between", formulae: [0, 1] },
    ]);

    ws.dataValidations.remove("B2:B10");
    expect(ws.dataValidations.find("B5")).toBeUndefined();
  });
});

describe("compat facade — cell.note bridges ExcelJS's flat texts array (COMMENTS-SPEC.md §3)", () => {
  it("a bare string passes straight through", () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("S");
    const cell = ws.getCell("A1") as unknown as { note: unknown };

    cell.note = "hello";
    expect(cell.note).toBe("hello");
  });

  it("a rich texts array round-trips through our RichText wrapper", () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("S");
    const cell = ws.getCell("A1") as unknown as { note: unknown };

    cell.note = { texts: [{ text: "bold", font: { bold: true } }, { text: " plain" }] };
    expect(cell.note).toEqual({
      texts: [{ text: "bold", font: { bold: true } }, { text: " plain" }],
    });
  });

  it("an author carries through get/set (a compat extension — ExcelJS itself has no author field)", () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("S");
    const cell = ws.getCell("A1") as unknown as { note: unknown };

    cell.note = { texts: "hi", author: "Alice" };
    expect(cell.note).toEqual({ texts: "hi", author: "Alice" });
  });
});

describe("compat facade — conditional formatting colour-shape bridge (CONDITIONAL-FORMATTING-SPEC §4)", () => {
  interface CfFacadeWorksheet {
    addConditionalFormatting: (cf: unknown) => unknown;
    conditionalFormattings: unknown[];
  }

  it("normalizes a colorScale's {argb} colours to bare ARGB strings before they reach the writer", async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("S") as unknown as CfFacadeWorksheet;

    ws.addConditionalFormatting({
      ref: "A1:A10",
      rules: [
        {
          type: "colorScale",
          cfvo: [{ type: "min" }, { type: "max" }],
          color: [{ argb: "FFF8696B" }, { argb: "FF63BE7B" }],
        },
      ],
    });

    // If the core model still held {argb} objects, the writer's `String(color)` would emit
    // "[object Object]" instead of an rgb attribute — so a real `rgb="FFF8696B"` in the written
    // bytes proves the objects were normalized to bare strings before serialization.
    const buf = await (
      wb as unknown as { xlsx: { writeBuffer: () => Promise<Uint8Array> } }
    ).xlsx.writeBuffer();
    const back = readWorkbookBuffer(buf);
    const data = back.getWorksheet("S");
    if (!data) throw new Error("S sheet missing");
    expect(data.conditionalFormattings).toEqual([
      {
        ref: "A1:A10",
        rules: [
          {
            type: "colorScale",
            cfvo: [{ type: "min" }, { type: "max" }],
            color: ["FFF8696B", "FF63BE7B"],
            priority: 1,
          },
        ],
      },
    ]);
  });

  it("normalizes a dataBar's single {argb} colour before it reaches the writer", async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("S") as unknown as CfFacadeWorksheet;

    ws.addConditionalFormatting({
      ref: "D1:D20",
      rules: [
        {
          type: "dataBar",
          cfvo: [{ type: "min" }, { type: "max" }],
          color: { argb: "FF638EC6" },
        },
      ],
    });

    const buf = await (
      wb as unknown as { xlsx: { writeBuffer: () => Promise<Uint8Array> } }
    ).xlsx.writeBuffer();
    const back = readWorkbookBuffer(buf);
    const data = back.getWorksheet("S");
    if (!data) throw new Error("S sheet missing");
    expect(data.conditionalFormattings[0]?.rules[0]).toMatchObject({ color: "FF638EC6" });
  });

  it("addConditionalFormatting's return value and the conditionalFormattings getter denormalize back to {argb}", () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("S") as unknown as CfFacadeWorksheet;

    const added = ws.addConditionalFormatting({
      ref: "A1:A10",
      rules: [
        {
          type: "colorScale",
          cfvo: [{ type: "min" }, { type: "max" }],
          color: [{ argb: "FFF8696B" }, { argb: "FF63BE7B" }],
        },
      ],
    }) as { rules: Array<{ color: unknown }> };

    expect(added.rules[0]?.color).toEqual([{ argb: "FFF8696B" }, { argb: "FF63BE7B" }]);
    expect(ws.conditionalFormattings).toEqual([
      {
        ref: "A1:A10",
        rules: [
          {
            type: "colorScale",
            cfvo: [{ type: "min" }, { type: "max" }],
            color: [{ argb: "FFF8696B" }, { argb: "FF63BE7B" }],
          },
        ],
      },
    ]);
  });

  it("cellIs and other dxf-bearing rules (already ExcelJS-shaped Style) pass through untouched", () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("S") as unknown as CfFacadeWorksheet;

    ws.addConditionalFormatting({
      ref: "A1:A10",
      rules: [
        {
          type: "cellIs",
          operator: "greaterThan",
          formulae: [5],
          style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFF8696B" } } },
        },
      ],
    });

    expect(ws.conditionalFormattings).toEqual([
      {
        ref: "A1:A10",
        rules: [
          {
            type: "cellIs",
            operator: "greaterThan",
            formulae: [5],
            style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFF8696B" } } },
          },
        ],
      },
    ]);
  });
});

describe("compat facade — eachRow/eachCell skip empty like ExcelJS (0.3.1)", () => {
  it("default iteration drops empty rows and explicitly-null cells", () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = "keep";
    ws.getCell("B1").value = null; // explicitly nulled -> skipped by default
    ws.getRow(2); // touched but valueless row -> skipped by default
    ws.getCell("A3").value = 7;

    const visitedRows: number[] = [];
    ws.eachRow((_row, n) => visitedRows.push(n));
    expect(visitedRows).toEqual([1, 3]); // row 2 skipped

    const cellValues: unknown[] = [];
    ws.getRow(1).eachCell((cell) => cellValues.push(cell.value));
    expect(cellValues).toEqual(["keep"]); // B1 (null) skipped
  });

  it("callback cells are still proxied (report ExcelJS ValueType numbers)", () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = "hi";
    let seenType: unknown;
    ws.getRow(1).eachCell((cell) => {
      seenType = cell.type;
    });
    expect(seenType).toBe(ValueType.String); // 3, not the core CellType number
  });

  it("includeEmpty:true still enumerates the empty rows/cells", () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = "x";
    ws.getCell("B1").value = null;
    ws.getRow(2);
    ws.getCell("A3").value = "y";

    const rows: number[] = [];
    ws.eachRow({ includeEmpty: true }, (_row, n) => rows.push(n));
    expect(rows).toEqual([1, 2, 3]); // row 2 included

    const cols: number[] = [];
    ws.getRow(1).eachCell({ includeEmpty: true }, (_cell, c) => cols.push(c));
    expect(cols).toEqual([1, 2]); // B1 included
  });
});
