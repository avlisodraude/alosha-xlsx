import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { PassThrough } from "node:stream";
import { describe, expect, it } from "vitest";
import type { CellValue, Workbook } from "../../../src/model/index.js";
import { Workbook as WorkbookModel } from "../../../src/model/index.js";
import { readWorkbookBuffer } from "../../../src/xlsx/read-workbook.js";
import { WorkbookWriter } from "../../../src/xlsx/stream/index.js";
import { writeWorkbookBuffer } from "../../../src/xlsx/write-workbook.js";
import { readZipArchive } from "../../../src/zip-layer.js";

// Fixed doc metadata so the streamed and buffered core-props parts are byte-identical (otherwise the
// `created`/`modified` timestamps would differ between two `new Workbook()`s).
const CREATED = new Date("2026-01-02T03:04:05.000Z");
const MODIFIED = new Date("2026-02-03T04:05:06.000Z");
const CREATOR = "Streaming E1";

/** A row plus optional per-cell number formats applied after the values are set. */
interface RowSpec {
  values: CellValue[];
  numFmt?: Array<[col: number, code: string]>;
}

/** A multi-type payload for "Data": string (+ duplicate for shared-string dedup), styled number,
 *  boolean, date, error, formula, rich text, and a hyperlink — every E1 cell kind. */
const DATA_ROWS: RowSpec[] = [
  { values: ["Hello", "Hello"] },
  { values: [42, true], numFmt: [[1, "0.00"]] },
  { values: [new Date("2026-01-15T00:00:00.000Z"), { error: "#REF!" }] },
  { values: [{ formula: "A2+1", result: 43 }] },
  {
    values: [
      { richText: [{ text: "bold", font: { bold: true } }, { text: " plain" }] },
      { text: "site", hyperlink: "https://example.com", tooltip: "open" },
    ],
  },
];

const SECOND_ROWS: RowSpec[] = [{ values: ["World", 7] }, { values: [false] }];

/** Populate a buffered model worksheet from row specs (row numbers assigned sequentially). */
function applyModel(ws: ReturnType<Workbook["addWorksheet"]>, specs: RowSpec[]): void {
  for (const spec of specs) {
    const row = ws.addRow(spec.values);
    for (const [col, code] of spec.numFmt ?? []) row.getCell(col).numFmt = code;
  }
}

/** Build the buffered reference workbook. */
function buildModel(): Workbook {
  const wb = new WorkbookModel();
  wb.creator = CREATOR;
  wb.created = CREATED;
  wb.modified = MODIFIED;
  applyModel(wb.addWorksheet("Data"), DATA_ROWS);
  applyModel(wb.addWorksheet("Second"), SECOND_ROWS);
  return wb;
}

/** Stream the identical payload and return the whole archive. */
async function buildStream(options?: {
  useSharedStrings?: boolean;
  useStyles?: boolean;
}): Promise<Uint8Array> {
  const wb = new WorkbookWriter(options);
  wb.creator = CREATOR;
  wb.created = CREATED;
  wb.modified = MODIFIED;

  const applyStream = (ws: ReturnType<WorkbookWriter["addWorksheet"]>, specs: RowSpec[]): void => {
    for (const spec of specs) {
      const row = ws.addRow(spec.values);
      for (const [col, code] of spec.numFmt ?? []) row.getCell(col).numFmt = code;
      row.commit();
    }
  };

  const data = wb.addWorksheet("Data");
  applyStream(data, DATA_ROWS);
  data.commit();

  const second = wb.addWorksheet("Second");
  applyStream(second, SECOND_ROWS);
  second.commit();

  await wb.commit();
  return collect(wb);
}

/** Drain an async iterable of chunks into one buffer. */
async function collect(source: AsyncIterable<Uint8Array>): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  let total = 0;
  for await (const chunk of source) {
    chunks.push(chunk);
    total += chunk.length;
  }
  const out = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    out.set(chunk, offset);
    offset += chunk.length;
  }
  return out;
}

/** Stable JSON of a parsed workbook model (Dates → ISO), for structural deep-diffing. */
function modelJson(wb: Workbook): string {
  return JSON.stringify(wb.model, (_k, v) => (v instanceof Date ? v.toISOString() : v), 2);
}

describe("WorkbookWriter — structural equality with the buffered writer", () => {
  it("streams a multi-sheet, multi-type workbook equal to writeWorkbookBuffer({useSharedStrings:false})", async () => {
    const streamed = await buildStream({ useSharedStrings: false });
    const buffered = writeWorkbookBuffer(buildModel(), { useSharedStrings: false });

    const fromStream = readWorkbookBuffer(streamed);
    const fromBuffer = readWorkbookBuffer(buffered);
    expect(modelJson(fromStream)).toEqual(modelJson(fromBuffer));
  });

  it("matches the buffered part set (inline strings: no sharedStrings.xml)", async () => {
    const entries = readZipArchive(await buildStream({ useSharedStrings: false }));
    expect(new Set(entries.keys())).toEqual(
      new Set([
        "[Content_Types].xml",
        "_rels/.rels",
        "docProps/core.xml",
        "docProps/app.xml",
        "xl/workbook.xml",
        "xl/_rels/workbook.xml.rels",
        "xl/styles.xml",
        "xl/theme/theme1.xml",
        "xl/worksheets/sheet1.xml",
        "xl/worksheets/sheet2.xml",
        // sheet1 carries a hyperlink → its relationships part
        "xl/worksheets/_rels/sheet1.xml.rels",
      ]),
    );
    const sheet1 = new TextDecoder().decode(entries.get("xl/worksheets/sheet1.xml"));
    expect(sheet1).toContain('t="inlineStr"');
    expect(sheet1).toContain("<is><t>Hello</t></is>");
  });
});

describe("WorkbookWriter — shared-strings mode", () => {
  it("interns strings and emits sharedStrings.xml, equal to the buffered shared-strings output", async () => {
    const streamed = await buildStream({ useSharedStrings: true });
    const entries = readZipArchive(streamed);
    expect(entries.has("xl/sharedStrings.xml")).toBe(true);

    const sheet1 = new TextDecoder().decode(entries.get("xl/worksheets/sheet1.xml"));
    // "Hello" in A1 and B1 dedup to the same shared index.
    expect(sheet1).toContain('<c r="A1" t="s"><v>0</v></c>');
    expect(sheet1).toContain('<c r="B1" t="s"><v>0</v></c>');

    const buffered = writeWorkbookBuffer(buildModel(), { useSharedStrings: true });
    expect(modelJson(readWorkbookBuffer(streamed))).toEqual(
      modelJson(readWorkbookBuffer(buffered)),
    );
  });
});

describe("WorkbookWriter — edge cases", () => {
  it("writes a valid archive for an empty sheet", async () => {
    const wb = new WorkbookWriter();
    const ws = wb.addWorksheet("Empty");
    ws.commit();
    await wb.commit();

    const bytes = await collect(wb);
    const parsed = readWorkbookBuffer(bytes);
    expect(parsed.worksheets).toHaveLength(1);
    expect(parsed.worksheets[0]?.name).toBe("Empty");
    expect(parsed.worksheets[0]?.rowCount).toBe(0);

    const sheet1 = new TextDecoder().decode(readZipArchive(bytes).get("xl/worksheets/sheet1.xml"));
    expect(sheet1).toContain("<sheetData></sheetData>");
  });

  it("yields a valid archive through the async-iterable sink (no filename/stream)", async () => {
    const wb = new WorkbookWriter();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRow(["a", 1]).commit();
    ws.addRow(["b", 2]).commit();
    ws.commit();
    await wb.commit();

    // The writer itself is the async iterable of chunks.
    const chunks: Uint8Array[] = [];
    for await (const chunk of wb) chunks.push(chunk);
    expect(chunks.length).toBeGreaterThan(0);

    let total = 0;
    for (const c of chunks) total += c.length;
    const bytes = new Uint8Array(total);
    let off = 0;
    for (const c of chunks) {
      bytes.set(c, off);
      off += c.length;
    }
    const parsed = readWorkbookBuffer(bytes);
    expect(parsed.worksheets[0]?.getCell("A1").value).toBe("a");
    expect(parsed.worksheets[0]?.getCell("B2").value).toBe(2);
  });

  it("pipes into a Node Writable sink", async () => {
    const sink = new PassThrough();
    const chunks: Uint8Array[] = [];
    sink.on("data", (c: Buffer) => chunks.push(new Uint8Array(c)));

    const wb = new WorkbookWriter({ stream: sink });
    const ws = wb.addWorksheet("Piped");
    ws.addRow(["x", 9]).commit();
    ws.commit();
    await wb.commit();

    let total = 0;
    for (const c of chunks) total += c.length;
    const bytes = new Uint8Array(total);
    let off = 0;
    for (const c of chunks) {
      bytes.set(c, off);
      off += c.length;
    }
    const parsed = readWorkbookBuffer(bytes);
    expect(parsed.worksheets[0]?.getCell("A1").value).toBe("x");
  });

  it("writes to a file via the Node fs adapter", async () => {
    const dir = await mkdtemp(join(tmpdir(), "alosha-stream-"));
    const path = join(dir, "out.xlsx");
    try {
      const wb = new WorkbookWriter({ filename: path });
      const ws = wb.addWorksheet("File");
      ws.addRow(["saved", 3.14]).commit();
      ws.commit();
      await wb.commit();

      const parsed = readWorkbookBuffer(new Uint8Array(await readFile(path)));
      expect(parsed.worksheets[0]?.getCell("A1").value).toBe("saved");
      expect(parsed.worksheets[0]?.getCell("B1").value).toBe(3.14);
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("rejects interleaving two open worksheets", () => {
    const wb = new WorkbookWriter();
    const a = wb.addWorksheet("A");
    const b = wb.addWorksheet("B");
    a.addRow([1]).commit(); // opens A's part
    expect(() => b.addRow([2]).commit()).toThrow(/commit it before writing/);
  });
});

describe("WorkbookWriter — column widths (<cols>)", () => {
  it("streams column widths/styles/hidden flags equal to the buffered writer", async () => {
    const columns = [{ width: 15 }, { width: 30, style: { font: { bold: true } } }];

    const model = new WorkbookModel();
    const bufferedWs = model.addWorksheet("Sheet1");
    bufferedWs.columns = columns;
    bufferedWs.addRow(["a", "b"]);
    const buffered = writeWorkbookBuffer(model);

    const wb = new WorkbookWriter();
    const ws = wb.addWorksheet("Sheet1");
    ws.columns = columns; // must be set before the first row commits
    ws.addRow(["a", "b"]).commit();
    ws.commit();
    await wb.commit();
    const streamed = await collect(wb);

    expect(modelJson(readWorkbookBuffer(streamed))).toEqual(
      modelJson(readWorkbookBuffer(buffered)),
    );

    const parsed = readWorkbookBuffer(streamed);
    const sheet = parsed.getWorksheet("Sheet1");
    expect(sheet?.getColumn(1).width).toBe(15);
    expect(sheet?.getColumn(2).width).toBe(30);
    expect(sheet?.getColumn(2).font?.bold).toBe(true);
  });

  it("supports getColumn for a single column's width/style, set before the first row commits", async () => {
    const wb = new WorkbookWriter();
    const ws = wb.addWorksheet("Sheet1");
    ws.getColumn(2).width = 42;
    ws.getColumn(2).hidden = true;
    ws.addRow(["x", "y"]).commit();
    ws.commit();
    await wb.commit();

    const parsed = readWorkbookBuffer(await collect(wb));
    const sheet = parsed.getWorksheet("Sheet1");
    expect(sheet?.getColumn(2).width).toBe(42);
    expect(sheet?.getColumn(2).hidden).toBe(true);
  });

  it("ignores a column set after the first row commits (same rule as views/panes)", async () => {
    const wb = new WorkbookWriter();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRow(["a"]).commit(); // opens the part; the prologue (incl. <cols>) is already written
    ws.getColumn(1).width = 99; // too late — has no effect on the emitted archive
    ws.commit();
    await wb.commit();

    const parsed = readWorkbookBuffer(await collect(wb));
    expect(parsed.getWorksheet("Sheet1")?.getColumn(1).width).toBeUndefined();
  });
});
