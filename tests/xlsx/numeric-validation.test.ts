import { describe, expect, it } from "vitest";
import {
  CellType,
  colLetterToNumber,
  colNumberToLetter,
  dateToSerial,
  encodeAddress,
  InvalidAddressError,
  InvalidCellValueError,
  InvalidRangeError,
  Range,
  readWorkbookBuffer,
  serialToDate,
  Workbook,
  WorkbookWriter,
  writeWorkbookBuffer,
} from "../../src/index.js";
import { readZipArchive, writeZipArchive } from "../../src/zip-layer.js";

function foreignSheet(body: string): Uint8Array {
  const wb = new Workbook();
  wb.addWorksheet("S");
  const parts = readZipArchive(writeWorkbookBuffer(wb));
  parts.set("xl/worksheets/sheet1.xml", new TextEncoder().encode(`<worksheet>${body}</worksheet>`));
  return writeZipArchive(parts);
}

describe("consistent coordinate validation", () => {
  it.each([
    0,
    -1,
    1.5,
    NaN,
    Infinity,
    1048577,
  ])("rejects invalid row %s across numeric APIs", (row) => {
    const ws = new Workbook().addWorksheet("S");
    for (const operation of [
      () => ws.getRow(row),
      () => ws.findRow(row),
      () => ws.getCell(row, 1),
      () => ws.findCell(row, 1),
      () => encodeAddress(row, 1),
      () => ws.getRows(row, 1),
    ]) {
      expect(operation).toThrow(InvalidAddressError);
    }
    expect(ws.rowCount).toBe(0);
    expect(() => new Range(row, 1, row, 1)).toThrow(InvalidRangeError);
  });

  it.each([
    0,
    -1,
    1.5,
    NaN,
    Infinity,
    16385,
  ])("rejects invalid column %s across numeric APIs", (col) => {
    const ws = new Workbook().addWorksheet("S");
    for (const operation of [
      () => ws.getColumn(col),
      () => ws.getCell(1, col),
      () => ws.findCell(1, col),
      () => encodeAddress(1, col),
      () => colNumberToLetter(col),
    ]) {
      expect(operation).toThrow(InvalidAddressError);
    }
    expect(ws.rowCount).toBe(0);
    expect(() => ws.getRow(1).getCell(col)).toThrow(InvalidAddressError);
    expect(() => ws.getRow(1).findCell(col)).toThrow(InvalidAddressError);
    expect(() => new Range(1, col, 1, col)).toThrow(InvalidRangeError);
  });

  it("accepts the last valid row/column and rejects oversized letters", () => {
    const ws = new Workbook().addWorksheet("S");
    ws.getCell(1048576, 16384).value = 1;
    expect(ws.getCell("XFD1048576").value).toBe(1);
    expect(new Range(1, 1, 1048576, 16384).br).toBe("XFD1048576");
    expect(() => colLetterToNumber("XFE")).toThrow(InvalidAddressError);
    expect(() => colLetterToNumber("AAAA")).toThrow(InvalidAddressError);
    const range = new Range("A1:B2");
    expect(() => range.expand(1, 1, 1048577, 2)).toThrow(InvalidRangeError);
    expect(range.range).toBe("A1:B2");
  });

  it("preflights a numeric row span before allocating rows", () => {
    const ws = new Workbook().addWorksheet("S");
    expect(() => ws.getRows(1048576, 2)).toThrow(InvalidAddressError);
    expect(ws.rowCount).toBe(0);
  });

  it("rejects grid-overflowing edits before moving existing cells", () => {
    const ws = new Workbook().addWorksheet("S");
    ws.getCell("XFD1048576").value = "last";
    expect(() => ws.spliceRows(1, 0, ["inserted"])).toThrow(InvalidRangeError);
    expect(() => ws.spliceColumns(1, 0, [])).toThrow(InvalidRangeError);
    expect(() => ws.getRow(1048576).splice(1, 0, "inserted")).toThrow(InvalidRangeError);
    expect(() => ws.duplicateRow(1048576, 1)).toThrow(InvalidRangeError);
    expect(ws.getCell("XFD1048576").value).toBe("last");
    expect(ws.findCell("A1")).toBeUndefined();
  });

  it("preflights shifted annotation ranges before editing a sheet", () => {
    const ws = new Workbook().addWorksheet("S");
    ws.getCell("A1").value = "keep";
    ws.dataValidations.add("A1048576", { type: "whole", formulae: [0] });
    expect(() => ws.spliceRows(1, 0, ["inserted"])).toThrow(InvalidRangeError);
    expect(ws.getCell("A1").value).toBe("keep");
    expect(ws.rowCount).toBe(1);
  });

  it("rejects invalid numeric inserts before deleting rows or cells", () => {
    const ws = new Workbook().addWorksheet("S");
    ws.addRow(["keep"]);
    expect(() => ws.addRow([NaN])).toThrow(InvalidCellValueError);
    expect(ws.rowCount).toBe(1);
    expect(() => ws.spliceRows(1, 1, [Infinity])).toThrow(InvalidCellValueError);
    expect(() => ws.spliceColumns(1, 1, [0, NaN])).toThrow(InvalidCellValueError);
    expect(() => ws.getRow(1).splice(1, 1, NaN)).toThrow(InvalidCellValueError);
    expect(ws.getCell("A1").value).toBe("keep");
  });

  it.each([
    '<cols><col min="1" max="16385"/></cols>',
    '<cols><col min="1.5" max="2"/></cols>',
    '<cols><col min="2" max="1"/></cols>',
    '<sheetData><row r="1048577"/></sheetData>',
    '<sheetData><row r="NaN"/></sheetData>',
    '<sheetData><row r="1"><c r="XFE1"><v>1</v></c></row></sheetData>',
    '<sheetData><row r="1"><c r="A2"><v>1</v></c></row></sheetData>',
  ])("rejects malformed coordinates on import: %s", (body) => {
    expect(() => readWorkbookBuffer(foreignSheet(body))).toThrow();
  });
});

describe("finite spreadsheet values", () => {
  it.each([
    NaN,
    Infinity,
    -Infinity,
    new Date("invalid"),
  ])("rejects %s and preserves the previous value", (value) => {
    const cell = new Workbook().addWorksheet("S").getCell("A1");
    cell.value = 42;
    expect(() => {
      cell.value = value;
    }).toThrow(InvalidCellValueError);
    expect(cell.value).toBe(42);
    expect(() => {
      cell.value = { formula: "1+1", result: value };
    }).toThrow(InvalidCellValueError);
    expect(cell.value).toBe(42);
  });

  it("allows finite numbers, textual NaN, and explicit Excel error values", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.addRow([
      0,
      -42.5,
      Number.MAX_VALUE,
      "NaN",
      { error: "#NUM!" },
      { formula: "1+1", result: 2 },
    ]);
    const parsed = readWorkbookBuffer(writeWorkbookBuffer(wb)).worksheets[0];
    expect(parsed?.getCell("A1").value).toBe(0);
    expect(parsed?.getCell("C1").value).toBe(Number.MAX_VALUE);
    expect(parsed?.getCell("D1").type).toBe(CellType.String);
    expect(parsed?.getCell("E1").value).toEqual({ error: "#NUM!" });
  });

  it("revalidates an assigned Date or formula object mutated before serialization", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    const date = new Date("2024-01-01");
    ws.getCell("A1").value = date;
    date.setTime(NaN);
    expect(() => writeWorkbookBuffer(wb)).toThrow(InvalidCellValueError);
    ws.getCell("A1").value = 1;
    const formula = { formula: "1+1", result: 2 };
    ws.getCell("A1").value = formula;
    formula.result = Infinity;
    expect(() => writeWorkbookBuffer(wb)).toThrow(InvalidCellValueError);
  });

  it.each([
    "NaN",
    "Infinity",
    "-Infinity",
    "not-a-number",
  ])("rejects non-finite XML numeric values: %s", (value) => {
    for (const formula of ["", "<f>1+1</f>"]) {
      const bytes = foreignSheet(
        `<sheetData><row r="1"><c r="A1">${formula}<v>${value}</v></c></row></sheetData>`,
      );
      expect(() => readWorkbookBuffer(bytes)).toThrow(InvalidCellValueError);
    }
  });

  it.each([false, true])("rejects invalid date conversions (1904=%s)", (date1904) => {
    expect(() => dateToSerial(new Date("invalid"), date1904)).toThrow(InvalidCellValueError);
    for (const serial of [NaN, Infinity, -Infinity, 1e100]) {
      expect(() => serialToDate(serial, date1904)).toThrow(InvalidCellValueError);
    }
  });
});

describe("streamed row recovery", () => {
  it("does not consume row numbers after invalid payloads or invalid column keys", async () => {
    const wb = new WorkbookWriter();
    const ws = wb.addWorksheet("S");
    expect(() => ws.addRow([{} as never])).toThrow(InvalidCellValueError);
    expect(() => ws.addRow([1, NaN])).toThrow(InvalidCellValueError);
    expect(() => ws.addRow({ XFE: 3 })).toThrow(InvalidAddressError);
    const first = ws.addRow([7]);
    expect(first.number).toBe(1);
    first.commit();
    expect(() => ws.addRow([Infinity])).toThrow(InvalidCellValueError);
    const second = ws.addRow([8]);
    expect(second.number).toBe(2);
    second.commit();
    await wb.commit();
    const chunks: Uint8Array[] = [];
    for await (const chunk of wb) chunks.push(chunk);
    const parsed = readWorkbookBuffer(Buffer.concat(chunks)).worksheets[0];
    expect(parsed?.rowCount).toBe(2);
    expect(parsed?.getCell("A1").value).toBe(7);
    expect(parsed?.getCell("A2").value).toBe(8);
  });

  it("permits correcting a mutated invalid result before retrying commit", async () => {
    const wb = new WorkbookWriter();
    const ws = wb.addWorksheet("S");
    const formula = { formula: "1+1", result: 2 };
    const row = ws.addRow([formula]);
    formula.result = NaN;
    expect(() => row.commit()).toThrow(InvalidCellValueError);
    formula.result = 2;
    row.commit();
    await wb.commit();
  });
});
