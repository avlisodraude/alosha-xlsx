import { describe, expect, it } from "vitest";
import { Workbook } from "../../src/model/index.js";
import { writeWorkbookBuffer } from "../../src/xlsx/write-workbook.js";
import { readZipArchive } from "../../src/zip-layer.js";

/** Decode a required archive part to text, failing loudly if it is absent. */
function part(entries: Map<string, Uint8Array>, path: string): string {
  const bytes = entries.get(path);
  if (!bytes) throw new Error(`missing part: ${path}`);
  return new TextDecoder().decode(bytes);
}

/** A small, deterministic fixture exercising every value type, a merge, a number format, a styled
 *  column, a formula, and a second sheet. Fixed dates keep the metadata parts stable. */
function fixture(): Workbook {
  const wb = new Workbook();
  wb.created = new Date("2026-01-02T03:04:05.000Z");
  wb.modified = new Date("2026-01-02T03:04:05.000Z");
  wb.creator = "Eduardo";

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

  const ws2 = wb.addWorksheet("Second");
  ws2.getCell("A1").value = "World";
  return wb;
}

const EXPECTED_PARTS = [
  "[Content_Types].xml",
  "_rels/.rels",
  "docProps/core.xml",
  "docProps/app.xml",
  "xl/workbook.xml",
  "xl/_rels/workbook.xml.rels",
  "xl/styles.xml",
  "xl/sharedStrings.xml",
  "xl/theme/theme1.xml",
  "xl/worksheets/sheet1.xml",
  "xl/worksheets/sheet2.xml",
];

describe("writeWorkbookBuffer — archive shape", () => {
  it("produces exactly the expected part set", () => {
    const entries = readZipArchive(writeWorkbookBuffer(fixture()));
    expect(new Set(entries.keys())).toEqual(new Set(EXPECTED_PARTS));
  });

  it("omits sharedStrings.xml when useSharedStrings is false", () => {
    const entries = readZipArchive(writeWorkbookBuffer(fixture(), { useSharedStrings: false }));
    expect(entries.has("xl/sharedStrings.xml")).toBe(false);
    const sheet1 = part(entries, "xl/worksheets/sheet1.xml");
    expect(sheet1).toContain('t="inlineStr"');
    expect(sheet1).toContain("<is><t>Hello</t></is>");
  });
});

describe("writeWorkbookBuffer — worksheet grid", () => {
  it("maps each value type to the right cell encoding", () => {
    const entries = readZipArchive(writeWorkbookBuffer(fixture()));
    const sheet1 = part(entries, "xl/worksheets/sheet1.xml");

    expect(sheet1).toContain('<dimension ref="A1:B5"/>');
    // string -> shared index (both A1 and B1 reference index 0)
    expect(sheet1).toContain('<c r="A1" t="s"><v>0</v></c>');
    expect(sheet1).toContain('<c r="B1" t="s"><v>0</v></c>');
    // number (t omitted) carrying the styled numFmt
    expect(sheet1).toMatch(/<c r="A2"[^>]*><v>42<\/v><\/c>/);
    // boolean / error / formula
    expect(sheet1).toContain('<c r="B2" t="b"><v>1</v></c>');
    expect(sheet1).toContain('<c r="B3" t="e"><v>#REF!</v></c>');
    expect(sheet1).toContain("<f>A2+1</f><v>43</v>");
    // merged column + merge region
    expect(sheet1).toContain('<col min="1" max="1" width="20" customWidth="1"');
    expect(sheet1).toContain('<mergeCells count="1"><mergeCell ref="A5:B5"/></mergeCells>');
  });

  it("interns strings once across sheets and reports counts", () => {
    const entries = readZipArchive(writeWorkbookBuffer(fixture()));
    const sst = part(entries, "xl/sharedStrings.xml");
    expect(sst).toContain("<t>Hello</t>");
    expect(sst).toContain("<t>World</t>");
    expect(sst).toContain("<t>Merged</t>");
    // Hello added twice, Merged once, World once => 4 references, 3 unique.
    expect(sst).toContain('count="4"');
    expect(sst).toContain('uniqueCount="3"');
  });
});

describe("writeWorkbookBuffer — hyperlinks wiring", () => {
  it("writes xl/worksheets/_rels/<file>.rels only for sheets that have hyperlinks", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Links");
    ws.getCell("A1").value = { text: "Anthropic", hyperlink: "https://anthropic.com" };
    wb.addWorksheet("Plain").getCell("A1").value = "no links here";

    const entries = readZipArchive(writeWorkbookBuffer(wb));
    expect(entries.has("xl/worksheets/_rels/sheet1.xml.rels")).toBe(true);
    expect(entries.has("xl/worksheets/_rels/sheet2.xml.rels")).toBe(false);

    const rels = part(entries, "xl/worksheets/_rels/sheet1.xml.rels");
    expect(rels).toContain('Target="https://anthropic.com" TargetMode="External"');
    const sheet1 = part(entries, "xl/worksheets/sheet1.xml");
    expect(sheet1).toContain('<hyperlink ref="A1" r:id="rId1"/>');
  });
});

describe("writeWorkbookBuffer — shared & array formulas", () => {
  it("promotes a plain-formula master + sharedFormula slaves into an si-numbered shared group", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Calc");
    ws.getCell("A1").value = 10;
    ws.getCell("A2").value = 20;
    ws.getCell("A3").value = 30;
    ws.getCell("B1").value = { formula: "A1*2", result: 20 };
    ws.getCell("B2").value = { sharedFormula: "B1", result: 40 };
    ws.getCell("B3").value = { sharedFormula: "B1", result: 60 };

    const sheet1 = part(readZipArchive(writeWorkbookBuffer(wb)), "xl/worksheets/sheet1.xml");
    // The master carries the expression, the group's bounding ref, and the shared index.
    expect(sheet1).toContain('<f t="shared" ref="B1:B3" si="0">A1*2</f>');
    // Slaves carry only the shared index — the expression is not duplicated.
    expect(sheet1).toContain('<c r="B2"><f t="shared" si="0"/><v>40</v></c>');
    expect(sheet1).toContain('<c r="B3"><f t="shared" si="0"/><v>60</v></c>');
  });

  it('emits an array formula as t="array" with its ref', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Arr");
    ws.getCell("C1").value = {
      formula: "SUM(A1:A3*B1:B3)",
      shareType: "array",
      ref: "C1:C1",
      result: 100,
    };

    const sheet1 = part(readZipArchive(writeWorkbookBuffer(wb)), "xl/worksheets/sheet1.xml");
    expect(sheet1).toContain('<f t="array" ref="C1:C1">SUM(A1:A3*B1:B3)</f>');
    expect(sheet1).toContain("<v>100</v>");
  });
});

describe("writeWorkbookBuffer — exceljs round-trip", () => {
  it("re-opens a shared + array formula group in exceljs with cached results intact", async (ctx) => {
    const spec = "exceljs";
    let ExcelJS: { Workbook: new () => unknown } | undefined;
    try {
      const mod = (await import(spec)) as { default?: unknown } & Record<string, unknown>;
      ExcelJS = (mod.default ?? mod) as { Workbook: new () => unknown };
    } catch {
      ctx.skip();
      return;
    }

    const wb = new Workbook();
    const ws = wb.addWorksheet("Calc");
    ws.getCell("A1").value = 10;
    ws.getCell("A2").value = 20;
    ws.getCell("A3").value = 30;
    ws.getCell("B1").value = { formula: "A1*2", result: 20 };
    ws.getCell("B2").value = { sharedFormula: "B1", result: 40 };
    ws.getCell("B3").value = { sharedFormula: "B1", result: 60 };
    ws.getCell("C1").value = {
      formula: "SUM(A1:A3)",
      shareType: "array",
      ref: "C1:C1",
      result: 60,
    };

    // biome-ignore lint/suspicious/noExplicitAny: exceljs has no static types available here.
    const loaded = new (ExcelJS.Workbook as any)();
    await loaded.xlsx.load(Buffer.from(writeWorkbookBuffer(wb)));
    const calc = loaded.getWorksheet("Calc");

    // exceljs re-opens the file without repair and resolves every cached result; the master's own
    // expression survives, and the shared group expands so a slave reports the master's formula too.
    expect(calc.getCell("B1").result).toBe(20);
    expect(calc.getCell("B1").formula).toBe("A1*2");
    expect(calc.getCell("B2").result).toBe(40);
    expect(calc.getCell("B3").result).toBe(60);
    expect(calc.getCell("C1").result).toBe(60);
  });

  it("re-opens in exceljs with values, merge and number format intact", async (ctx) => {
    // Variable specifier: not statically resolved, so the suite still typechecks/runs when the
    // optional devDependency is absent (it simply skips).
    const spec = "exceljs";
    let ExcelJS: { Workbook: new () => unknown } | undefined;
    try {
      const mod = (await import(spec)) as { default?: unknown } & Record<string, unknown>;
      ExcelJS = (mod.default ?? mod) as { Workbook: new () => unknown };
    } catch {
      ctx.skip();
      return;
    }

    const buf = writeWorkbookBuffer(fixture());
    // biome-ignore lint/suspicious/noExplicitAny: exceljs has no static types available here.
    const loaded = new (ExcelJS.Workbook as any)();
    await loaded.xlsx.load(Buffer.from(buf));

    const ws = loaded.getWorksheet("Data");
    expect(ws.getCell("A1").value).toBe("Hello");
    expect(ws.getCell("A2").value).toBe(42);
    expect(ws.getCell("B2").value).toBe(true);
    expect(ws.getCell("A2").numFmt).toBe("0.00");
    expect(loaded.getWorksheet("Second").getCell("A1").value).toBe("World");
    expect(ws.getCell("A5").isMerged).toBe(true);
  });
});
