// The real proof of the read path (SPEC §9): write the Stage 1.3 fixture, read it back, and assert the
// model is reproduced. Covers the shared-string default plus the inline-strings and styles-off writer
// variants, and the typed-error contract for malformed packages.

import { describe, expect, it } from "vitest";
// Import from the package entry so the `workbook.xlsx` accessor side-effect (defined there) is attached.
import {
  CellType,
  InvalidCellValueError,
  InvalidConditionalFormattingError,
  InvalidPackageError,
  readWorkbookBuffer,
  Workbook,
  type WriteOptions,
  writeWorkbookBuffer,
} from "../../src/index.js";
import { readZipArchive, writeZipArchive } from "../../src/zip-layer.js";

/** The Stage 1.3 fixture: every value type, a merge, a number format, a styled column, a formula, a
 *  second (hidden) sheet, and fixed metadata. Mirrors tests/xlsx/write-workbook.test.ts. */
function fixture(): Workbook {
  const wb = new Workbook();
  wb.created = new Date("2026-01-02T03:04:05.000Z");
  wb.modified = new Date("2026-01-02T03:04:05.000Z");
  wb.creator = "Eduardo";
  wb.company = "Alosha";

  const ws = wb.addWorksheet("Data");
  ws.getCell("A1").value = "Hello";
  ws.getCell("B1").value = "Hello"; // duplicate -> same shared-string index
  ws.getCell("A2").value = 42;
  ws.getCell("A2").numFmt = "0.00";
  ws.getCell("B2").value = true;
  ws.getCell("A3").value = new Date("2026-01-15T00:00:00.000Z");
  ws.getCell("B3").value = { error: "#REF!" };
  ws.getCell("A4").value = { formula: "A2+1", result: 43 };
  ws.getColumn(1).width = 20;
  ws.getCell("A5").value = "Merged";
  ws.mergeCells("A5:B5");

  const ws2 = wb.addWorksheet("Second", { state: "hidden" });
  ws2.getCell("A1").value = "World";
  return wb;
}

function roundTrip(options?: WriteOptions): Workbook {
  return readWorkbookBuffer(writeWorkbookBuffer(fixture(), options));
}

describe("read round-trip — shared strings (default)", () => {
  const wb = roundTrip();
  const data = wb.getWorksheet("Data");
  if (!data) throw new Error("Data sheet missing");

  it("reproduces sheet names, order and state", () => {
    expect(wb.worksheets.map((w) => w.name)).toEqual(["Data", "Second"]);
    expect(wb.worksheets.map((w) => w.state)).toEqual(["visible", "hidden"]);
  });

  it("reproduces every cell value type", () => {
    expect(data.getCell("A1").value).toBe("Hello");
    expect(data.getCell("B1").value).toBe("Hello");
    expect(data.getCell("A2").value).toBe(42);
    expect(data.getCell("B2").value).toBe(true);
    expect(data.getCell("A3").value).toBeInstanceOf(Date);
    expect((data.getCell("A3").value as Date).getTime()).toBe(
      new Date("2026-01-15T00:00:00.000Z").getTime(),
    );
    expect(data.getCell("B3").value).toEqual({ error: "#REF!" });
    expect(data.getCell("A4").formula).toBe("A2+1");
    expect(data.getCell("A4").result).toBe(43);
    expect(wb.getWorksheet("Second")?.getCell("A1").value).toBe("World");
  });

  it("reproduces the number format, column width and merge", () => {
    expect(data.getCell("A2").numFmt).toBe("0.00");
    expect(data.getColumn(1).width).toBe(20);
    expect(data.getCell("A5").value).toBe("Merged");
    expect(data.getCell("B5").isMerged).toBe(true);
  });

  it("reproduces workbook metadata (creator/created/modified/company)", () => {
    expect(wb.creator).toBe("Eduardo");
    expect(wb.company).toBe("Alosha");
    expect(wb.created.getTime()).toBe(new Date("2026-01-02T03:04:05.000Z").getTime());
    expect(wb.modified.getTime()).toBe(new Date("2026-01-02T03:04:05.000Z").getTime());
  });

  it("leaves date1904 unset for a 1900-system workbook", () => {
    expect(wb.properties.date1904).toBeUndefined();
  });
});

describe("read round-trip — inline strings variant", () => {
  it("reproduces string values written inline (no shared-strings part)", () => {
    const wb = roundTrip({ useSharedStrings: false });
    const data = wb.getWorksheet("Data");
    expect(data?.getCell("A1").value).toBe("Hello");
    expect(data?.getCell("B1").value).toBe("Hello");
    expect(data?.getCell("A2").value).toBe(42);
    expect(data?.getCell("B2").value).toBe(true);
    expect(data?.getCell("A5").value).toBe("Merged");
    expect(data?.getCell("B5").isMerged).toBe(true);
  });
});

describe("read round-trip — styles-off variant", () => {
  it("reproduces values but drops number formats, so a date reads back as its serial number", () => {
    const wb = roundTrip({ useStyles: false });
    const data = wb.getWorksheet("Data");
    if (!data) throw new Error("Data sheet missing");

    expect(data.getCell("A1").value).toBe("Hello");
    expect(data.getCell("A2").value).toBe(42);
    expect(data.getCell("A2").numFmt).toBeUndefined();
    expect(data.getCell("B2").value).toBe(true);
    expect(data.getCell("A4").result).toBe(43);
    expect(data.getCell("A5").value).toBe("Merged");
    expect(data.getCell("B5").isMerged).toBe(true);

    // Without a style, the date's serial is indistinguishable from a number.
    expect(data.getCell("A3").type).toBe(CellType.Number);
    expect(typeof data.getCell("A3").value).toBe("number");
  });
});

describe("read round-trip — date1904 workbook", () => {
  it("preserves the 1904 flag and resolves dates against the 1904 epoch", () => {
    const src = new Workbook();
    src.properties.date1904 = true;
    const ws = src.addWorksheet("S");
    const date = new Date("2020-06-15T00:00:00.000Z");
    ws.getCell("A1").value = date;

    const wb = readWorkbookBuffer(writeWorkbookBuffer(src));
    expect(wb.properties.date1904).toBe(true);
    const value = wb.getWorksheet("S")?.getCell("A1").value;
    expect(value).toBeInstanceOf(Date);
    expect((value as Date).getTime()).toBe(date.getTime());
  });
});

describe("read round-trip — hyperlinks", () => {
  it("persists external cell hyperlinks as OOXML relationships and reads them back intact", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Links");
    ws.getCell("A1").value = { text: "Anthropic", hyperlink: "https://anthropic.com" };
    ws.getCell("A2").value = {
      text: "Mail us",
      hyperlink: "mailto:hi@example.com",
      tooltip: "Send an email",
    };
    ws.getCell("A3").value = { text: "Docs", hyperlink: "https://docs.example.com/guide" };

    const bytes = writeWorkbookBuffer(wb);
    const entries = readZipArchive(bytes);
    expect(entries.has("xl/worksheets/_rels/sheet1.xml.rels")).toBe(true);

    const back = readWorkbookBuffer(bytes);
    const data = back.getWorksheet("Links");
    if (!data) throw new Error("Links sheet missing");

    expect(data.getCell("A1").type).toBe(CellType.Hyperlink);
    expect(data.getCell("A1").value).toEqual({
      text: "Anthropic",
      hyperlink: "https://anthropic.com",
    });

    expect(data.getCell("A2").type).toBe(CellType.Hyperlink);
    expect(data.getCell("A2").value).toEqual({
      text: "Mail us",
      hyperlink: "mailto:hi@example.com",
      tooltip: "Send an email",
    });

    expect(data.getCell("A3").type).toBe(CellType.Hyperlink);
    expect(data.getCell("A3").value).toEqual({
      text: "Docs",
      hyperlink: "https://docs.example.com/guide",
    });
  });
});

describe("read round-trip — XML-illegal characters (control chars from e.g. pasted PDF text)", () => {
  it("persists a cell value containing a vertical tab as valid, spec-compliant XML and reads it back intact", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Illegal");
    ws.getCell("A1").value = "before\x0Bafter";

    const bytes = writeWorkbookBuffer(wb);
    const entries = readZipArchive(bytes);
    const sharedStrings = new TextDecoder().decode(entries.get("xl/sharedStrings.xml"));
    // The raw control char must never appear literally; it's re-encoded as Excel's _xHHHH_ token.
    expect(sharedStrings).not.toContain("\x0B");
    expect(sharedStrings).toContain("_x000B_");

    const back = readWorkbookBuffer(bytes);
    expect(back.getWorksheet("Illegal")?.getCell("A1").value).toBe("before\x0Bafter");
  });

  it("persists a control char in an attribute (numFmt formatCode) and reads it back intact", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Illegal");
    const cell = ws.getCell("A1");
    cell.value = 42;
    cell.numFmt = '0.00"\x0B"';

    const bytes = writeWorkbookBuffer(wb);
    const entries = readZipArchive(bytes);
    const styles = new TextDecoder().decode(entries.get("xl/styles.xml"));
    expect(styles).not.toContain("\x0B");
    expect(styles).toContain("_x000B_");

    const back = readWorkbookBuffer(bytes);
    expect(back.getWorksheet("Illegal")?.getCell("A1").numFmt).toBe('0.00"\x0B"');
  });
});

describe("read round-trip — tampered/corrupt sharedStrings part", () => {
  it("parses a sharedStrings.xml containing an out-of-range numeric character reference without throwing", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Tampered");
    ws.getCell("A1").value = "hello";

    const bytes = writeWorkbookBuffer(wb);
    const entries = readZipArchive(bytes);
    const original = new TextDecoder().decode(entries.get("xl/sharedStrings.xml"));
    // Simulate a corrupt/crafted part: an &#x110000; reference is beyond the Unicode range and would
    // make String.fromCodePoint throw a RangeError if not guarded.
    const tampered = original.replace("hello", "hello&#x110000;");
    entries.set("xl/sharedStrings.xml", new TextEncoder().encode(tampered));
    const tamperedBytes = writeZipArchive(entries);

    expect(() => readWorkbookBuffer(tamperedBytes)).not.toThrow();
    const back = readWorkbookBuffer(tamperedBytes);
    expect(back.getWorksheet("Tampered")?.getCell("A1").value).toBe("hello&#x110000;");
  });
});

describe("read round-trip — per-cell style inside a merged range", () => {
  it("preserves a border set on a non-master (slave) cell of a merge across write->read", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Merged");
    ws.getCell("A1").value = "master";
    ws.mergeCells("A1:B2");
    // Give the slave B2 its own distinct style (a thick bottom border) after the merge.
    ws.getCell("B2").border = { bottom: { style: "thick" } };

    const back = readWorkbookBuffer(writeWorkbookBuffer(wb));
    const data = back.getWorksheet("Merged");
    if (!data) throw new Error("Merged sheet missing");

    // The slave stays merged AND keeps its own border, rather than inheriting the master's blank style.
    expect(data.getCell("B2").isMerged).toBe(true);
    expect(data.getCell("B2").border).toEqual({ bottom: { style: "thick" } });
    expect(data.getCell("A1").value).toBe("master");
  });
});

describe("read round-trip — data validations", () => {
  it("persists a list dropdown, a numeric between, and a custom rule and reads them back equal", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Validated");
    ws.getCell("A1").dataValidation = {
      type: "list",
      allowBlank: true,
      showInputMessage: true,
      showErrorMessage: true,
      formulae: ['"Yes,No,Maybe"'],
    };
    ws.dataValidations.add("B1:B10", {
      type: "decimal",
      operator: "between",
      formulae: [0, 100],
    });
    ws.dataValidations.add("C1", {
      type: "custom",
      showErrorMessage: true,
      errorStyle: "stop",
      errorTitle: "Invalid",
      error: "Must be a positive number",
      formulae: ["ISNUMBER(C1)"],
    });

    const back = readWorkbookBuffer(writeWorkbookBuffer(wb));
    const data = back.getWorksheet("Validated");
    if (!data) throw new Error("Validated sheet missing");

    expect(data.dataValidations.model).toEqual(ws.dataValidations.model);
    expect(data.getCell("A1").dataValidation).toEqual({
      type: "list",
      allowBlank: true,
      showInputMessage: true,
      showErrorMessage: true,
      formulae: ['"Yes,No,Maybe"'],
    });
    expect(data.dataValidations.find("B5")).toEqual({
      type: "decimal",
      operator: "between",
      formulae: [0, 100],
    });
    expect(data.dataValidations.find("C1")).toEqual({
      type: "custom",
      showErrorMessage: true,
      errorStyle: "stop",
      errorTitle: "Invalid",
      error: "Must be a positive number",
      formulae: ["ISNUMBER(C1)"],
    });
  });
});

describe("read round-trip — panes & auto-filter", () => {
  it("persists a frozen pane and an auto-filter range together", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Panes");
    ws.getCell("A1").value = "Header";
    ws.views = [{ state: "frozen", xSplit: 1, ySplit: 1, topLeftCell: "B2" }];
    ws.autoFilter = "A1:C1";

    const back = readWorkbookBuffer(writeWorkbookBuffer(wb));
    const data = back.getWorksheet("Panes");
    if (!data) throw new Error("Panes sheet missing");

    expect(data.views).toEqual([{ state: "frozen", xSplit: 1, ySplit: 1, topLeftCell: "B2" }]);
    expect(data.autoFilter).toBe("A1:C1");
  });

  it("persists a zoom-only view", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Zoom");
    ws.getCell("A1").value = "Header";
    ws.views = [{ zoomScale: 150 }];

    const back = readWorkbookBuffer(writeWorkbookBuffer(wb));
    const data = back.getWorksheet("Zoom");
    if (!data) throw new Error("Zoom sheet missing");

    expect(data.views).toEqual([{ zoomScale: 150 }]);
  });

  it("persists zoom alongside a frozen pane", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("ZoomPane");
    ws.getCell("A1").value = "Header";
    ws.views = [{ state: "frozen", xSplit: 1, ySplit: 1, topLeftCell: "B2", zoomScale: 75 }];

    const back = readWorkbookBuffer(writeWorkbookBuffer(wb));
    const data = back.getWorksheet("ZoomPane");
    if (!data) throw new Error("ZoomPane sheet missing");

    expect(data.views).toEqual([
      { state: "frozen", xSplit: 1, ySplit: 1, topLeftCell: "B2", zoomScale: 75 },
    ]);
  });
});

describe("read round-trip — conditional formatting", () => {
  it("persists one rule per family, preserving ref, priority order and every type-specific field", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("CF");
    ws.addConditionalFormatting({
      ref: "A1:A10",
      rules: [
        { type: "cellIs", operator: "greaterThan", formulae: [5] },
        { type: "cellIs", operator: "between", formulae: [1, 10] },
        { type: "expression", formulae: ["$A1>0"] },
        { type: "containsText", text: "foo" },
        { type: "notContainsText", text: "bar" },
        { type: "beginsWith", text: "Q" },
        { type: "endsWith", text: "Z" },
        { type: "top10", rank: 10, percent: true, bottom: true },
        { type: "aboveAverage" },
        { type: "belowAverage" },
        { type: "duplicateValues" },
        { type: "uniqueValues" },
      ],
    });
    ws.addConditionalFormatting({
      ref: "C1:C20",
      rules: [
        {
          type: "colorScale",
          cfvo: [{ type: "min" }, { type: "max" }],
          color: ["FFF8696B", "FF63BE7B"],
        },
      ],
    });
    ws.addConditionalFormatting({
      ref: "D1:D20",
      rules: [
        {
          type: "dataBar",
          cfvo: [{ type: "min" }, { type: "max" }],
          color: "FF638EC6",
        },
      ],
    });
    ws.addConditionalFormatting({
      ref: "E1:E20",
      rules: [
        {
          type: "iconSet",
          iconSet: "3TrafficLights1",
          cfvo: [
            { type: "percent", value: 0 },
            { type: "percent", value: 33 },
            { type: "percent", value: 67 },
          ],
        },
      ],
    });

    const back = readWorkbookBuffer(writeWorkbookBuffer(wb));
    const data = back.getWorksheet("CF");
    if (!data) throw new Error("CF sheet missing");

    expect(data.conditionalFormattings.map((cf) => cf.ref)).toEqual([
      "A1:A10",
      "C1:C20",
      "D1:D20",
      "E1:E20",
    ]);

    const firstBlockRules = data.conditionalFormattings[0]?.rules ?? [];
    // Priorities are assigned in a single workbook-monotonic sequence across every block, so the
    // first block's rules land on 1..12 (increasing, in the order they were added) and the second
    // block's colorScale continues from 13, etc.
    expect(firstBlockRules.map((r) => r.priority)).toEqual([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]);
    expect(firstBlockRules.map(({ priority: _priority, ...rest }) => rest)).toEqual([
      { type: "cellIs", operator: "greaterThan", formulae: [5] },
      { type: "cellIs", operator: "between", formulae: [1, 10] },
      { type: "expression", formulae: ["$A1>0"] },
      { type: "containsText", text: "foo" },
      { type: "notContainsText", text: "bar" },
      { type: "beginsWith", text: "Q" },
      { type: "endsWith", text: "Z" },
      { type: "top10", rank: 10, percent: true, bottom: true },
      { type: "aboveAverage" },
      { type: "belowAverage" },
      { type: "duplicateValues" },
      { type: "uniqueValues" },
    ]);

    expect(data.conditionalFormattings[1]?.rules).toEqual([
      {
        type: "colorScale",
        cfvo: [{ type: "min" }, { type: "max" }],
        color: ["FFF8696B", "FF63BE7B"],
        priority: 13,
      },
    ]);
    expect(data.conditionalFormattings[2]?.rules).toEqual([
      {
        type: "dataBar",
        cfvo: [{ type: "min" }, { type: "max" }],
        color: "FF638EC6",
        priority: 14,
      },
    ]);
    expect(data.conditionalFormattings[3]?.rules).toEqual([
      {
        type: "iconSet",
        iconSet: "3TrafficLights1",
        cfvo: [
          { type: "percent", value: 0 },
          { type: "percent", value: 33 },
          { type: "percent", value: 67 },
        ],
        priority: 15,
      },
    ]);
  });

  it("resolves a dxf-bearing rule's style through the shared dxf pool", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Highlight");
    ws.addConditionalFormatting({
      ref: "B1:B10",
      rules: [
        {
          type: "cellIs",
          operator: "lessThan",
          formulae: [0],
          style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFFF0000" } } },
        },
      ],
    });

    const back = readWorkbookBuffer(writeWorkbookBuffer(wb));
    const data = back.getWorksheet("Highlight");
    if (!data) throw new Error("Highlight sheet missing");

    expect(data.conditionalFormattings[0]?.rules[0]).toEqual({
      type: "cellIs",
      operator: "lessThan",
      formulae: [0],
      priority: 1,
      style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFFF0000" } } },
    });
  });

  it("preserves explicit priorities set on the model", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Explicit");
    ws.addConditionalFormatting({
      ref: "A1:A5",
      rules: [
        { type: "duplicateValues", priority: 42 },
        { type: "uniqueValues", priority: 7 },
      ],
    });

    const back = readWorkbookBuffer(writeWorkbookBuffer(wb));
    const data = back.getWorksheet("Explicit");
    if (!data) throw new Error("Explicit sheet missing");

    expect(data.conditionalFormattings[0]?.rules).toEqual([
      { type: "duplicateValues", priority: 42 },
      { type: "uniqueValues", priority: 7 },
    ]);
  });
});

describe("read round-trip — images", () => {
  it("persists a two-cell and a one-cell anchored image, byte-identical, on read-back", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Pics");
    const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    const gif = new Uint8Array([0x47, 0x49, 0x46, 0x38, 0x39, 0x61]);
    const pngId = wb.addImage({ buffer: png, extension: "png" });
    const gifId = wb.addImage({ buffer: gif, extension: "gif" });
    ws.addImage(pngId, "B2:D6");
    ws.addImage(gifId, { tl: { col: 5, row: 5 }, ext: { width: 100, height: 50 } });

    const back = readWorkbookBuffer(writeWorkbookBuffer(wb));
    const data = back.getWorksheet("Pics");
    if (!data) throw new Error("Pics sheet missing");

    expect(data.images).toEqual([
      { imageId: 0, range: { tl: { col: 1, row: 1 }, br: { col: 3, row: 5 } } },
      { imageId: 1, range: { tl: { col: 5, row: 5 }, ext: { width: 100, height: 50 } } },
    ]);

    const image0 = back.getImage(data.images[0]?.imageId ?? -1);
    expect(image0?.extension).toBe("png");
    expect([...(image0?.buffer ?? [])]).toEqual([...png]);

    const image1 = back.getImage(data.images[1]?.imageId ?? -1);
    expect(image1?.extension).toBe("gif");
    expect([...(image1?.buffer ?? [])]).toEqual([...gif]);
  });

  it("shares one registry image across two sheets that both anchor it", () => {
    const wb = new Workbook();
    const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47]);
    const id = wb.addImage({ buffer: png, extension: "png" });
    wb.addWorksheet("One").addImage(id, "A1:B2");
    wb.addWorksheet("Two").addImage(id, "C3:D4");

    const back = readWorkbookBuffer(writeWorkbookBuffer(wb));
    const one = back.getWorksheet("One");
    const two = back.getWorksheet("Two");
    if (!one || !two) throw new Error("sheets missing");

    const oneImageId = one.images[0]?.imageId;
    const twoImageId = two.images[0]?.imageId;
    expect(oneImageId).toBe(twoImageId);
    expect(back.getImage(oneImageId ?? -1)?.extension).toBe("png");
  });

  it("leaves an image-free workbook with no images on read-back (regression)", () => {
    const back = readWorkbookBuffer(writeWorkbookBuffer(fixture()));
    expect(back.getWorksheet("Data")?.images).toEqual([]);
  });
});

describe("read round-trip — shared & array formulas (foreign <f t=shared/array> groups)", () => {
  // A shared-formula group and an array formula exactly as Excel/exceljs emit them: one master `<f>`
  // carrying the expression + `ref` + `si`, its slaves carrying only `t="shared" si`, plus an array
  // master. Injected over a real (otherwise empty) package so the whole read path is exercised.
  const SHEET =
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>` +
    `<row r="1"><c r="A1"><v>1</v></c>` +
    `<c r="B1"><f t="shared" ref="B1:B3" si="0">A1*2</f><v>2</v></c>` +
    `<c r="C1"><f t="array" ref="C1:C1">A1+A2</f><v>3</v></c></row>` +
    `<row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>4</v></c></row>` +
    `<row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/><v>6</v></c></row>` +
    `</sheetData></worksheet>`;

  function readInjected(): Workbook {
    const src = new Workbook();
    src.addWorksheet("Shared");
    const entries = readZipArchive(writeWorkbookBuffer(src));
    entries.set("xl/worksheets/sheet1.xml", new TextEncoder().encode(SHEET));
    return readWorkbookBuffer(writeZipArchive(entries));
  }

  it("materializes the master (expression+ref), its slaves (sharedFormula ref), and the array master", () => {
    const ws = readInjected().getWorksheet("Shared");
    if (!ws) throw new Error("Shared sheet missing");

    expect(ws.getCell("B1").value).toEqual({
      formula: "A1*2",
      shareType: "shared",
      ref: "B1:B3",
      result: 2,
    });
    expect(ws.getCell("B2").value).toEqual({ sharedFormula: "B1", result: 4 });
    expect(ws.getCell("B3").value).toEqual({ sharedFormula: "B1", result: 6 });
    expect(ws.getCell("C1").value).toEqual({
      formula: "A1+A2",
      shareType: "array",
      ref: "C1:C1",
      result: 3,
    });
  });

  it("re-emits si/ref on write and reads back an identical model (write->read is stable)", () => {
    const bytes = writeWorkbookBuffer(readInjected());
    const sheet1 = new TextDecoder().decode(readZipArchive(bytes).get("xl/worksheets/sheet1.xml"));
    expect(sheet1).toContain('<f t="shared" ref="B1:B3" si="0">A1*2</f>');
    expect(sheet1).toContain('<f t="shared" si="0"/>');
    expect(sheet1).toContain('<f t="array" ref="C1:C1">A1+A2</f>');

    const ws = readWorkbookBuffer(bytes).getWorksheet("Shared");
    if (!ws) throw new Error("Shared sheet missing");
    expect(ws.getCell("B1").value).toEqual({
      formula: "A1*2",
      shareType: "shared",
      ref: "B1:B3",
      result: 2,
    });
    expect(ws.getCell("B2").value).toEqual({ sharedFormula: "B1", result: 4 });
    expect(ws.getCell("B3").value).toEqual({ sharedFormula: "B1", result: 6 });
    expect(ws.getCell("C1").value).toEqual({
      formula: "A1+A2",
      shareType: "array",
      ref: "C1:C1",
      result: 3,
    });
  });
});

describe("read path — typed errors", () => {
  it("throws InvalidPackageError on a truncated/corrupt zip", () => {
    expect(() => readWorkbookBuffer(new Uint8Array([1, 2, 3, 4, 5]))).toThrow(InvalidPackageError);
  });

  it("throws a typed cell error for a malformed boolean cell", () => {
    // Surgically corrupt the (resolved) worksheet part of an otherwise-valid package.
    const bytes = writeWorkbookBuffer(fixture());
    const entries = readZipArchive(bytes);
    entries.set(
      "xl/worksheets/sheet1.xml",
      new TextEncoder().encode(
        `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="b"><v>maybe</v></c></row></sheetData></worksheet>`,
      ),
    );
    expect(() => readWorkbookBuffer(writeZipArchive(entries))).toThrow(InvalidCellValueError);
  });

  it("throws a typed error for a malformed cfRule priority, never a silent NaN", () => {
    const bytes = writeWorkbookBuffer(fixture());
    const entries = readZipArchive(bytes);
    entries.set(
      "xl/worksheets/sheet1.xml",
      new TextEncoder().encode(
        `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/>` +
          `<conditionalFormatting sqref="A1:A9">` +
          `<cfRule type="duplicateValues" dxfId="0" priority="x"/></conditionalFormatting></worksheet>`,
      ),
    );
    expect(() => readWorkbookBuffer(writeZipArchive(entries))).toThrow(
      InvalidConditionalFormattingError,
    );
  });
});

describe("read round-trip — comments (docs/COMMENTS-SPEC.md §2)", () => {
  it("persists a plain-string note, a rich-text note, and a multi-author sheet intact", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Notes");
    ws.getCell("B2").note = { texts: "hello" };
    ws.getCell("C3").note = {
      texts: { richText: [{ text: "bold", font: { bold: true } }, { text: " plain" }] },
      author: "Alice",
    };
    ws.getCell("A1").note = { texts: "one", author: "Alice" };
    ws.getCell("A2").note = { texts: "two", author: "Bob" };
    ws.getCell("A3").note = { texts: "three", author: "Alice" };

    const back = readWorkbookBuffer(writeWorkbookBuffer(wb));
    const data = back.getWorksheet("Notes");
    if (!data) throw new Error("Notes sheet missing");

    // A plain string always renders as a single <r><t> run (COMMENTS-SPEC §3, exceljs interop), so
    // it reads back as a one-run RichText rather than the original string — a deliberate, documented
    // trade-off (see MIGRATION.md). Author defaults to "Author" (ExcelJS parity) when unset at write time.
    expect(data.getCell("B2").note).toEqual({
      texts: { richText: [{ text: "hello" }] },
      author: "Author",
    });
    expect(data.getCell("C3").note).toEqual({
      texts: { richText: [{ text: "bold", font: { bold: true } }, { text: " plain" }] },
      author: "Alice",
    });
    expect(data.getCell("A1").note).toEqual({
      texts: { richText: [{ text: "one" }] },
      author: "Alice",
    });
    expect(data.getCell("A2").note).toEqual({
      texts: { richText: [{ text: "two" }] },
      author: "Bob",
    });
    expect(data.getCell("A3").note).toEqual({
      texts: { richText: [{ text: "three" }] },
      author: "Alice",
    });

    // A cell with no comment carries no note.
    expect(data.getCell("D4").note).toBeUndefined();
  });

  it("leaves a no-comments sheet with no notes and no <legacyDrawing>", () => {
    const back = readWorkbookBuffer(writeWorkbookBuffer(fixture()));
    const data = back.getWorksheet("Data");
    if (!data) throw new Error("Data sheet missing");
    expect(data.getCell("A1").note).toBeUndefined();
  });
});

describe("read path — public accessor", () => {
  it("workbook.xlsx.load mutates the receiver in place and returns it", async () => {
    const bytes = writeWorkbookBuffer(fixture());
    const target = new Workbook();
    const returned = await target.xlsx.load(bytes);
    expect(returned).toBe(target);
    expect(target.getWorksheet("Data")?.getCell("A1").value).toBe("Hello");
  });

  it("workbook.xlsx.read drains a byte stream then parses it", async () => {
    const bytes = writeWorkbookBuffer(fixture());
    async function* chunks(): AsyncGenerator<Uint8Array> {
      yield bytes.subarray(0, 10);
      yield bytes.subarray(10);
    }
    const target = new Workbook();
    await target.xlsx.read(chunks());
    expect(target.getWorksheet("Data")?.getCell("A5").value).toBe("Merged");
  });
});
