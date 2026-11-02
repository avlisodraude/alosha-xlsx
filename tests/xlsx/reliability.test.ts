import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";
import {
  readWorkbookBuffer,
  Workbook,
  WorkbookWriter,
  writeWorkbookBuffer,
} from "../../src/index.js";
import { readZipArchive } from "../../src/zip-layer.js";

async function streamBytes(writer: WorkbookWriter): Promise<Uint8Array> {
  await writer.commit();
  const chunks: Uint8Array[] = [];
  for await (const chunk of writer) chunks.push(chunk);
  return Buffer.concat(chunks);
}

describe("serialization reliability", () => {
  it.each([
    true,
    false,
  ])("preserves literal escapes, rich text, notes, and formats (shared=%s)", (useSharedStrings) => {
    const wb = new Workbook();
    wb.creator = "_x0041_";
    const ws = wb.addWorksheet("_x0041_");
    const literal = "_x0041_ _x005F_ _x000B_\r\n\x0B & < 😀";
    ws.getCell("A1").value = literal;
    ws.getCell("B1").value = {
      richText: [{ text: literal, font: { bold: true, name: "_x0041_" } }],
    };
    ws.getCell("C1").value = { formula: '"_x0041_"', result: "_x0041_" };
    ws.getCell("D1").value = { text: "_x0041_", hyperlink: "https://example.com/_x0041_" };
    ws.getCell("A1").numFmt = '"_x0041_"0';
    ws.getCell("A1").note = { texts: literal, author: "_x0041_" };
    const bytes = writeWorkbookBuffer(wb, { useSharedStrings });
    const parsed = readWorkbookBuffer(bytes);
    const sheet = parsed.getWorksheet("_x0041_");
    expect(parsed.creator).toBe("_x0041_");
    expect(sheet?.getCell("A1").value).toBe(literal);
    expect(sheet?.getCell("B1").value).toEqual(ws.getCell("B1").value);
    expect(sheet?.getCell("C1").value).toEqual(ws.getCell("C1").value);
    expect(sheet?.getCell("D1").value).toEqual(ws.getCell("D1").value);
    expect(sheet?.getCell("A1").numFmt).toBe('"_x0041_"0');
    expect(sheet?.getCell("A1").note).toEqual({
      texts: { richText: [{ text: literal }] },
      author: "_x0041_",
    });
    const archive = readZipArchive(bytes);
    expect(new TextDecoder().decode(archive.get("xl/worksheets/_rels/sheet1.xml.rels"))).toContain(
      'Target="https://example.com/_x0041_"',
    );
  });

  it("preserves worksheet defaults through the workbook accessor", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Defaults");
    ws.properties.defaultRowHeight = 30;
    ws.properties.defaultColWidth = 25;
    ws.properties.tabColor = { argb: "FFFF0000" };
    ws.properties.outlineLevelRow = 2;
    ws.properties.outlineLevelCol = 1;
    ws.properties.dyDescent = 0.3;
    const parsed = new Workbook();
    await parsed.xlsx.load(await wb.xlsx.writeBuffer());
    expect(parsed.worksheets[0]?.properties.defaultRowHeight).toBe(30);
    expect(parsed.worksheets[0]?.properties.defaultColWidth).toBe(25);
    expect(parsed.worksheets[0]?.properties).toEqual(ws.properties);
    const streamed = new WorkbookWriter();
    const streamedSheet = streamed.addWorksheet("Defaults");
    Object.assign(streamedSheet.properties, ws.properties);
    const streamedParsed = readWorkbookBuffer(await streamBytes(streamed));
    expect(streamedParsed.worksheets[0]?.properties).toEqual(ws.properties);
  });

  it("streams rich text plus explicit shared and array formula groups readable by ExcelJS", async () => {
    const wb = new WorkbookWriter();
    const ws = wb.addWorksheet("S");
    const rich = { richText: [{ text: "_x0041_", font: { bold: true } }, { text: " plain" }] };
    ws.addRow([
      { formula: "B1*2", result: 6, shareType: "shared", ref: "A1:A2" },
      3,
      rich,
    ]).commit();
    ws.addRow([{ sharedFormula: "A1", result: 8 }, 4]).commit();
    ws.addRow([{ formula: "B3:B4*2", shareType: "array", ref: "A3:A4", result: 10 }, 5]).commit();
    ws.addRow([12, 6]).commit();
    const bytes = await streamBytes(wb);
    const parsed = readWorkbookBuffer(bytes).getWorksheet("S");
    expect(parsed?.getCell("A2").value).toEqual({ sharedFormula: "A1", result: 8 });
    expect(parsed?.getCell("A3").value).toEqual({
      formula: "B3:B4*2",
      shareType: "array",
      ref: "A3:A4",
      result: 10,
    });
    expect(parsed?.getCell("C1").value).toEqual(rich);
    const foreign = new ExcelJS.Workbook();
    // ExcelJS's legacy Buffer declaration predates the generic Node 22 Buffer type.
    await foreign.xlsx.load(
      Buffer.from(bytes) as unknown as Parameters<typeof foreign.xlsx.load>[0],
    );
    expect(foreign.getWorksheet("S")?.getCell("A2").formula).toBe("B2*2");
    expect(foreign.getWorksheet("S")?.getCell("C1").value).toEqual(rich);
  });

  it("rejects undeclared shared masters so no empty formula is emitted", async () => {
    const wb = new WorkbookWriter();
    const ws = wb.addWorksheet("S");
    ws.addRow([{ formula: "1+1", result: 2 }]).commit();
    const row = ws.addRow([{ sharedFormula: "A1", result: 2 }]);
    expect(() => row.commit()).toThrow(/preceding master/);
    row.getCell(1).value = { formula: "1+1", result: 2 };
    row.commit();
    const parsed = readWorkbookBuffer(await streamBytes(wb));
    expect(parsed.worksheets[0]?.getCell("A2").formula).toBe("1+1");
  });
});

describe("streaming lifecycle", () => {
  it.each([
    "",
    "bad/name",
    "x".repeat(32),
    "History",
    "'edge",
  ])("rejects invalid name %j before consuming a sheet id", (name) => {
    const wb = new WorkbookWriter();
    expect(() => wb.addWorksheet(name)).toThrow();
    expect(wb.addWorksheet().name).toBe("Sheet1");
  });

  it("rejects repeated/out-of-order commits and uncommitted rows without damaging the archive", async () => {
    const wb = new WorkbookWriter();
    const ws = wb.addWorksheet("S");
    const a = ws.addRow([1]);
    const b = ws.addRow([2]);
    expect(() => b.commit()).toThrow(/creation order/);
    expect(() => ws.commit()).toThrow(/all rows/);
    a.commit();
    expect(() => a.commit()).toThrow(/already committed/);
    b.commit();
    const bytes = await streamBytes(wb);
    const xml = new TextDecoder().decode(readZipArchive(bytes).get("xl/worksheets/sheet1.xml"));
    expect(xml.match(/<row /g)).toHaveLength(2);
    expect(readWorkbookBuffer(bytes).worksheets[0]?.getCell("A2").value).toBe(2);
  });

  it("recovers a second sheet after an attempted interleave", async () => {
    const wb = new WorkbookWriter();
    const a = wb.addWorksheet("A");
    const b = wb.addWorksheet("B");
    a.addRow([1]).commit();
    const row = b.addRow([2]);
    expect(() => row.commit()).toThrow(/still open/);
    a.commit();
    row.commit();
    const parsed = readWorkbookBuffer(await streamBytes(wb));
    expect(parsed.getWorksheet("B")?.getCell("A1").value).toBe(2);
  });

  it("auto-commits an active later sheet before an earlier empty sheet", async () => {
    const wb = new WorkbookWriter();
    wb.addWorksheet("Empty");
    wb.addWorksheet("Data").addRow([42]).commit();
    const parsed = readWorkbookBuffer(await streamBytes(wb));
    expect(parsed.worksheets.map((sheet) => sheet.name)).toEqual(["Empty", "Data"]);
    expect(parsed.getWorksheet("Data")?.getCell("A1").value).toBe(42);
  });
});
