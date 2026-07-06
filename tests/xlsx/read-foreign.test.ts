// Foreign-producer proof (SPEC §9): build a workbook with ExcelJS, have it serialize the bytes, then
// parse them with our reader. This exercises producer quirks our own writer never emits (differing part
// names, `t="str"`, shared `<f>` groups, whitespace runs). Guarded on the optional `exceljs` devDep — a
// variable import specifier keeps the suite typechecking and running when it is absent (it just skips).

import { describe, expect, it } from "vitest";
import { readWorkbookBuffer } from "../../src/xlsx/read-workbook.js";

// biome-ignore lint/suspicious/noExplicitAny: exceljs ships no static types we depend on here.
type AnyExcelJS = { Workbook: new () => any };

async function loadExcelJS(): Promise<AnyExcelJS | undefined> {
  const spec = "exceljs";
  try {
    const mod = (await import(spec)) as { default?: unknown } & Record<string, unknown>;
    return (mod.default ?? mod) as AnyExcelJS;
  } catch {
    return undefined;
  }
}

describe("read foreign — exceljs-produced buffer", () => {
  it("parses values, a merge and a number format from an exceljs workbook", async (ctx) => {
    const ExcelJS = await loadExcelJS();
    if (!ExcelJS) {
      ctx.skip();
      return;
    }

    const src = new ExcelJS.Workbook();
    const sheet = src.addWorksheet("Foreign");
    sheet.getCell("A1").value = "hello";
    sheet.getCell("A2").value = 12.5;
    sheet.getCell("A2").numFmt = "0.00";
    sheet.getCell("B1").value = true;
    sheet.getCell("C1").value = { formula: "A2*2", result: 25 };
    sheet.getCell("A3").value = "top";
    sheet.mergeCells("A3:B3");

    const buffer = await src.xlsx.writeBuffer();
    const bytes = new Uint8Array(buffer as ArrayBuffer);

    const wb = readWorkbookBuffer(bytes);
    const ws = wb.getWorksheet("Foreign");
    expect(ws).toBeDefined();
    if (!ws) return;

    expect(ws.getCell("A1").value).toBe("hello");
    expect(ws.getCell("A2").value).toBe(12.5);
    expect(ws.getCell("A2").numFmt).toBe("0.00");
    expect(ws.getCell("B1").value).toBe(true);
    expect(ws.getCell("C1").formula).toBe("A2*2");
    expect(ws.getCell("C1").result).toBe(25);
    expect(ws.getCell("A3").value).toBe("top");
    expect(ws.getCell("B3").isMerged).toBe(true);
  });
});
