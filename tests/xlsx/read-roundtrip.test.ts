// The real proof of the read path (SPEC §9): write the Stage 1.3 fixture, read it back, and assert the
// model is reproduced. Covers the shared-string default plus the inline-strings and styles-off writer
// variants, and the typed-error contract for malformed packages.

import { describe, expect, it } from "vitest";
// Import from the package entry so the `workbook.xlsx` accessor side-effect (defined there) is attached.
import {
  CellType,
  InvalidCellValueError,
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
