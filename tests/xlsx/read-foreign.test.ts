// Foreign-producer proof (SPEC §9): build a workbook with ExcelJS, have it serialize the bytes, then
// parse them with our reader. This exercises producer quirks our own writer never emits (differing part
// names, `t="str"`, shared `<f>` groups, whitespace runs). Guarded on the optional `exceljs` devDep — a
// variable import specifier keeps the suite typechecking and running when it is absent (it just skips).

import { describe, expect, it } from "vitest";
import { Workbook } from "../../src/model/index.js";
import { readWorkbookBuffer } from "../../src/xlsx/read-workbook.js";
import { writeWorkbookBuffer } from "../../src/xlsx/write-workbook.js";

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

describe("read foreign — hyperlinks", () => {
  it("recovers a hyperlink cell written by exceljs", async (ctx) => {
    const ExcelJS = await loadExcelJS();
    if (!ExcelJS) {
      ctx.skip();
      return;
    }

    const src = new ExcelJS.Workbook();
    const sheet = src.addWorksheet("Links");
    sheet.getCell("A1").value = { text: "Anthropic", hyperlink: "https://anthropic.com" };
    sheet.getCell("A2").value = {
      text: "Mail us",
      hyperlink: "mailto:hi@example.com",
      tooltip: "Send an email",
    };

    const buffer = await src.xlsx.writeBuffer();
    const bytes = new Uint8Array(buffer as ArrayBuffer);

    const wb = readWorkbookBuffer(bytes);
    const ws = wb.getWorksheet("Links");
    expect(ws).toBeDefined();
    if (!ws) return;

    expect(ws.getCell("A1").value).toEqual({
      text: "Anthropic",
      hyperlink: "https://anthropic.com",
    });
    expect(ws.getCell("A2").value).toEqual({
      text: "Mail us",
      hyperlink: "mailto:hi@example.com",
      tooltip: "Send an email",
    });
  });

  it("writes a hyperlink that re-opens in exceljs with the link intact", async (ctx) => {
    const ExcelJS = await loadExcelJS();
    if (!ExcelJS) {
      ctx.skip();
      return;
    }

    const wb = new Workbook();
    const ws = wb.addWorksheet("Links");
    ws.getCell("A1").value = { text: "Anthropic", hyperlink: "https://anthropic.com" };

    const bytes = writeWorkbookBuffer(wb);
    const loaded = new ExcelJS.Workbook();
    await loaded.xlsx.load(Buffer.from(bytes));

    const cell = loaded.getWorksheet("Links").getCell("A1");
    expect(cell.value).toEqual({ text: "Anthropic", hyperlink: "https://anthropic.com" });
  });
});

describe("read foreign — data validations", () => {
  it("recovers a list dropdown written by exceljs", async (ctx) => {
    const ExcelJS = await loadExcelJS();
    if (!ExcelJS) {
      ctx.skip();
      return;
    }

    const src = new ExcelJS.Workbook();
    const sheet = src.addWorksheet("Validated");
    sheet.getCell("A1").dataValidation = {
      type: "list",
      allowBlank: true,
      formulae: ['"Yes,No,Maybe"'],
    };
    // ExcelJS omits the `operator` attribute for its schema default ("between"); our reader must
    // resolve it the same way exceljs itself does on read.
    sheet.getCell("B1").dataValidation = {
      type: "decimal",
      operator: "between",
      formulae: [0, 100],
    };

    const buffer = await src.xlsx.writeBuffer();
    const bytes = new Uint8Array(buffer as ArrayBuffer);

    const wb = readWorkbookBuffer(bytes);
    const ws = wb.getWorksheet("Validated");
    expect(ws).toBeDefined();
    if (!ws) return;

    expect(ws.getCell("A1").dataValidation).toEqual({
      type: "list",
      allowBlank: true,
      formulae: ['"Yes,No,Maybe"'],
    });
    expect(ws.dataValidations.find("B1")).toEqual({
      type: "decimal",
      operator: "between",
      formulae: [0, 100],
    });
  });

  it("writes a dropdown that re-opens in exceljs with the constraint intact", async (ctx) => {
    const ExcelJS = await loadExcelJS();
    if (!ExcelJS) {
      ctx.skip();
      return;
    }

    const wb = new Workbook();
    const ws = wb.addWorksheet("Validated");
    ws.getCell("A1").dataValidation = {
      type: "list",
      allowBlank: true,
      formulae: ['"Yes,No,Maybe"'],
    };

    const bytes = writeWorkbookBuffer(wb);
    const loaded = new ExcelJS.Workbook();
    await loaded.xlsx.load(Buffer.from(bytes));

    const cell = loaded.getWorksheet("Validated").getCell("A1");
    expect(cell.dataValidation).toEqual({
      type: "list",
      allowBlank: true,
      formulae: ['"Yes,No,Maybe"'],
    });
  });
});
