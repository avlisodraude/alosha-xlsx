import { describe, expect, it, vi } from "vitest";
import {
  DEFAULT_READ_LIMITS,
  Range,
  ReadLimitError,
  readWorkbookBuffer,
  Workbook,
  Worksheet,
  writeWorkbookBuffer,
} from "../../src/index.js";
import { readZipArchive, writeZipArchive } from "../../src/zip-layer.js";

function packageWithSheets(...bodies: string[]): Uint8Array {
  const wb = new Workbook();
  for (let i = 0; i < bodies.length; i++) wb.addWorksheet(`Sheet${i + 1}`);
  const parts = readZipArchive(writeWorkbookBuffer(wb));
  for (const [i, body] of bodies.entries()) {
    parts.set(
      `xl/worksheets/sheet${i + 1}.xml`,
      new TextEncoder().encode(
        `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${body}</worksheet>`,
      ),
    );
  }
  return writeZipArchive(parts);
}

const cells =
  '<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData>';
const merge = '<mergeCells><mergeCell ref="A1:B2"/></mergeCells>';
const columns = '<cols><col min="1" max="2" width="10"/></cols>';

describe("workbook model import budgets", () => {
  it("checks merged-area size before enumerating a huge range", () => {
    const bytes = packageWithSheets('<mergeCells><mergeCell ref="A1:XFD1048576"/></mergeCells>');
    const enumerate = vi.spyOn(Range.prototype, "forEachAddress");
    try {
      expect(() => readWorkbookBuffer(bytes)).toThrow(ReadLimitError);
      expect(enumerate).not.toHaveBeenCalled();
      expect(() => readWorkbookBuffer(bytes)).toThrow(/maxMergeCells/);
    } finally {
      enumerate.mockRestore();
    }
  });

  it("counts actual cells including merge slaves without double-counting existing cells", () => {
    const bytes = packageWithSheets(cells + merge);
    const parsed = readWorkbookBuffer(bytes, { maxCells: 4, maxRows: 2, maxMergeCells: 4 });
    expect(parsed.worksheets[0]?.model.rows.reduce((n, row) => n + row.cells.length, 0)).toBe(4);
    expect(() => readWorkbookBuffer(bytes, { maxCells: 3, maxMergeCells: Infinity })).toThrow(
      /maxCells/,
    );
  });

  it.each([
    ["maxCells", cells, 3, 4],
    ["maxRows", '<sheetData><row r="1"/><row r="1048576"/></sheetData>', 3, 4],
    ["maxMergeCells", merge, 7, 8],
    ["maxColumnDefinitions", columns, 3, 4],
  ] as const)("shares %s across worksheets and accepts its exact boundary", (limit, body, below, exact) => {
    const bytes = packageWithSheets(body, body);
    expect(() => readWorkbookBuffer(bytes, { [limit]: below })).toThrow(
      new ReadLimitError(limit, below, exact),
    );
    expect(readWorkbookBuffer(bytes, { [limit]: exact }).worksheets).toHaveLength(2);
    expect(readWorkbookBuffer(bytes, { [limit]: Infinity }).worksheets).toHaveLength(2);
  });

  it("rejects expanded column definitions before creating columns", () => {
    const bytes = packageWithSheets('<cols><col min="1" max="16384" width="10"/></cols>');
    const getColumn = vi.spyOn(Worksheet.prototype, "getColumn");
    try {
      expect(() => readWorkbookBuffer(bytes, { maxColumnDefinitions: 10 })).toThrow(
        /maxColumnDefinitions/,
      );
      expect(getColumn).not.toHaveBeenCalled();
    } finally {
      getColumn.mockRestore();
    }
  });

  it("counts repeated column ranges as expansion work", () => {
    const bytes = packageWithSheets('<cols><col min="1" max="2"/><col min="1" max="2"/></cols>');
    expect(() => readWorkbookBuffer(bytes, { maxColumnDefinitions: 3 })).toThrow(
      /maxColumnDefinitions/,
    );
  });

  it("counts cells materialized only by comments", () => {
    const wb = new Workbook();
    wb.addWorksheet("S").getCell("D5").note = { texts: "note only" };
    expect(() => readWorkbookBuffer(writeWorkbookBuffer(wb), { maxCells: 0 })).toThrow(/maxCells/);
  });

  it("counts allocated rows rather than sparse row positions", () => {
    const bytes = packageWithSheets('<sheetData><row r="1048576"/></sheetData>');
    expect(readWorkbookBuffer(bytes, { maxRows: 1 }).worksheets[0]?.rowCount).toBe(1048576);
    expect(() => readWorkbookBuffer(bytes, { maxRows: 0 })).toThrow(/maxRows/);
  });

  it("resets budgets on every read and removes them before returning the model", () => {
    const bytes = packageWithSheets(cells);
    for (let i = 0; i < 2; i++) {
      const wb = readWorkbookBuffer(bytes, { maxCells: 2, maxRows: 1 });
      const ws = wb.worksheets[0];
      if (!ws) throw new Error("missing worksheet");
      ws.getCell("C2").value = 3;
      expect(ws.getCell("C2").value).toBe(3);
    }
  });

  it("preserves the receiving workbook when a model budget fails", async () => {
    const receiver = new Workbook();
    receiver.addWorksheet("Original").getCell("A1").value = 7;
    await expect(
      receiver.xlsx.load(packageWithSheets(merge), { maxCells: 3 }),
    ).rejects.toMatchObject({ code: "READ_LIMIT_EXCEEDED", limit: "maxCells" });
    expect(receiver.worksheets[0]?.name).toBe("Original");
    expect(receiver.worksheets[0]?.getCell("A1").value).toBe(7);
  });

  it.each([
    "maxCells",
    "maxRows",
    "maxMergeCells",
    "maxColumnDefinitions",
  ] as const)("validates %s options", (limit) => {
    const bytes = packageWithSheets(cells);
    expect(() => readWorkbookBuffer(bytes, { [limit]: -1 })).toThrow(RangeError);
    expect(() => readWorkbookBuffer(bytes, { [limit]: 0.5 })).toThrow(RangeError);
    expect(DEFAULT_READ_LIMITS[limit]).toBeGreaterThan(0);
  });
});
