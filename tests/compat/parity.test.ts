// The parity corpus — the real proof of the compat shim (docs/COMPAT-SPEC.md §8, integration plan).
//
// Each corpus entry is ONE `build(wb)` function run twice: once against the real `exceljs` package,
// once against `@alosha/xlsx/compat`'s `ExcelJS.Workbook`. Both are the same ExcelJS surface, so the
// same code populates either engine unmodified — that sameness is the point. We then:
//
//   1. Load BOTH resulting buffers back with real `exceljs` and diff a fixed set of sampled cells
//      (value/type/numFmt/two style flags), each sheet's merges, and the sheet-name order. Cross-
//      loading through exceljs normalizes away trivially-different-but-equivalent XML, so any
//      remaining diff is a genuine behavioural divergence.
//   2. Load the compat-produced buffer with our OWN `readWorkbookBuffer` and diff it against the same
//      exceljs-loaded reference — proving the shim's output round-trips through our own reader too.
//
// Guarded on the optional `exceljs` devDep via a variable import specifier (keeps the suite
// typechecking/running when it's absent — it just skips, mirroring read-foreign.test.ts /
// write-workbook.test.ts).
//
// Deliberately NOT covered here: hyperlink cell values. `{ text, hyperlink }` round-trips as plain
// display text through our writer (a known, pre-existing limitation of the core library — see the
// README's "Known limitations" and docs/MIGRATION.md), so a real exceljs-written hyperlink and a
// compat-written one are genuinely NOT equivalent after reload (the former keeps its relationship and
// reports `ValueType.Hyperlink`; the latter degrades to a plain string). Asserting equality here would
// either be flaky or paper over a real gap, so it's called out as a documented divergence instead.

import { describe, expect, it } from "vitest";
import { toExcelJSValueType } from "../../src/compat/value-type.js";
import { readWorkbookBuffer } from "../../src/xlsx/read-workbook.js";

// exceljs ships no static types we depend on, and the specifier below is intentionally not statically
// resolved (a plain string literal would make `import()` eager and break typechecking/running when the
// optional devDep is absent).
// biome-ignore lint/suspicious/noExplicitAny: exceljs/compat share a structural surface, not a static one.
type AnyWorkbook = any;
type WorkbookCtor = new () => AnyWorkbook;
type Build = (wb: AnyWorkbook) => void;

interface CorpusCase {
  name: string;
  build: Build;
  /** Sheet name -> the fixed set of cell addresses to sample ("sampled cell values" per §8, not a
   *  full-sheet enumeration — `eachRow`/`eachCell`'s default `includeEmpty:false` filtering differs
   *  subtly between the core model and ExcelJS, so explicit addresses sidestep that entirely). */
  sample: Record<string, string[]>;
}

// --- the corpus (§8: "~8-12 representative ExcelJS snippets") -----------------------------------

const corpus: CorpusCase[] = [
  {
    name: "every supported value type",
    build(wb) {
      const ws = wb.addWorksheet("Types");
      ws.getCell("A1").value = "hello";
      ws.getCell("A2").value = 42.5;
      ws.getCell("A3").value = true;
      ws.getCell("A4").value = false;
      ws.getCell("A5").value = new Date("2020-06-15T00:00:00.000Z");
      ws.getCell("A6").value = null;
      ws.getCell("A7").value = { error: "#DIV/0!" };
      ws.getCell("A8").value = { richText: [{ text: "rich" }, { text: "text" }] };
      ws.getCell("A9").value = { formula: "A2*2", result: 85 };
    },
    sample: { Types: ["A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9"] },
  },
  {
    name: "styled header row (font + fill)",
    build(wb) {
      const ws = wb.addWorksheet("Sheet1");
      const header = ws.getRow(1);
      header.values = ["Name", "Age"];
      header.font = { bold: true };
      header.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDDDDDD" } };
      ws.addRow(["Ada", 36]);
    },
    sample: { Sheet1: ["A1", "B1", "A2", "B2"] },
  },
  {
    name: "number formats (currency, percentage, date)",
    build(wb) {
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = 1234.5;
      ws.getCell("A1").numFmt = "#,##0.00";
      ws.getCell("A2").value = 0.4567;
      ws.getCell("A2").numFmt = "0.00%";
      ws.getCell("A3").value = new Date("2021-03-04T00:00:00.000Z");
      ws.getCell("A3").numFmt = "yyyy-mm-dd";
    },
    sample: { Sheet1: ["A1", "A2", "A3"] },
  },
  {
    name: "a merge",
    build(wb) {
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = "Merged Header";
      ws.mergeCells("A1:C1");
      ws.getCell("A2").value = "below";
    },
    sample: { Sheet1: ["A1", "B1", "C1", "A2"] },
  },
  {
    name: "multiple worksheets, order preserved",
    build(wb) {
      wb.addWorksheet("Alpha").getCell("A1").value = "alpha";
      wb.addWorksheet("Beta").getCell("A1").value = "beta";
      wb.addWorksheet("Gamma").getCell("A1").value = "gamma";
    },
    sample: { Alpha: ["A1"], Beta: ["A1"], Gamma: ["A1"] },
  },
  {
    name: "column keys + addRow(object)",
    build(wb) {
      const ws = wb.addWorksheet("Sheet1");
      ws.columns = [
        { header: "Name", key: "name", width: 20 },
        { header: "Age", key: "age", width: 10 },
      ];
      ws.addRow({ name: "Ada", age: 36 });
      ws.addRow({ name: "Grace", age: 85 });
    },
    sample: { Sheet1: ["A1", "B1", "A2", "B2", "A3", "B3"] },
  },
  {
    name: "formula with cached result",
    build(wb) {
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = 10;
      ws.getCell("A2").value = 20;
      ws.getCell("A3").value = { formula: "A1+A2", result: 30 };
      ws.getCell("A4").value = { formula: 'CONCATENATE("n=",A1)', result: "n=10" };
    },
    sample: { Sheet1: ["A1", "A2", "A3", "A4"] },
  },
  {
    name: "date1904 workbook",
    build(wb) {
      wb.properties.date1904 = true;
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = new Date("1990-07-04T00:00:00.000Z");
      ws.getCell("A2").value = new Date("2026-07-06T00:00:00.000Z");
    },
    sample: { Sheet1: ["A1", "A2"] },
  },
  {
    name: "row-level style inherited, then a cell overrides font (D2 whole-object precedence)",
    build(wb) {
      const ws = wb.addWorksheet("Sheet1");
      const row = ws.getRow(1);
      row.font = { name: "Arial", size: 12, italic: true };
      row.numFmt = "0.00";
      ws.getCell("A1").value = 1.5; // inherits the row's font + numFmt untouched
      const b1 = ws.getCell("B1");
      b1.value = 2.5;
      b1.font = { bold: true }; // whole-object replace: name/size/italic must NOT survive
    },
    sample: { Sheet1: ["A1", "B1"] },
  },
];

// --- sampling ------------------------------------------------------------------------------------

interface Sample {
  type: number;
  value: unknown;
  numFmt: string | null;
  /** Whether the resolved font is bold — a cheap, drift-resistant proxy for "styling survived",
   *  without deep-equality on nested color/theme objects that both engines are free to normalize
   *  differently (see docs/COMPAT-SPEC.md §8's DoD: values/type/numFmt/merges/structure, not full
   *  style objects). */
  bold: boolean;
  fillPattern: string | null;
}

function normalizeCellValue(v: unknown): unknown {
  if (v instanceof Date) return v.toISOString();
  if (v !== null && typeof v === "object") {
    const obj = v as Record<string, unknown>;
    if ("formula" in obj) return { formula: obj.formula, result: normalizeCellValue(obj.result) };
    if ("richText" in obj) {
      return { richText: (obj.richText as Array<{ text: string }>).map((r) => r.text) };
    }
    if ("error" in obj) return { error: obj.error };
    if ("text" in obj && "hyperlink" in obj) return { text: obj.text };
  }
  return v;
}

/** `pattern: "none"` and "no fill at all" both mean "unstyled" — but our reader leaves an unstyled
 *  cell's `fill` as `undefined` (fillId 0 is treated as "not set"), while ExcelJS always resolves
 *  the baseline `{pattern:"none"}` fill object. Canonicalize both to `null` so that pre-existing
 *  reader quirk (orthogonal to compat parity) doesn't make an otherwise-solid comparison flaky. */
function normalizedFillPattern(fill: { pattern?: string } | undefined): string | null {
  const p = fill?.pattern;
  return p === undefined || p === "none" ? null : p;
}

function sampleCell(cell: AnyWorkbook): Sample {
  const font = cell.font as { bold?: boolean } | undefined;
  return {
    type: cell.type,
    value: normalizeCellValue(cell.value),
    numFmt: cell.numFmt ?? null,
    bold: Boolean(font?.bold),
    fillPattern: normalizedFillPattern(cell.fill),
  };
}

/** For a workbook loaded through our OWN reader: its raw cells report core `CellType` numbers, not
 *  ExcelJS `ValueType` numbers, so the snapshot maps them through the same bridge the facade uses. */
function sampleCoreCell(cell: AnyWorkbook): Sample {
  return { ...sampleCell(cell), type: toExcelJSValueType(cell.type) };
}

interface Snapshot {
  sheetNames: string[];
  sheets: Record<string, { cells: Record<string, Sample>; merges: string[] }>;
}

function snapshot(
  wb: AnyWorkbook,
  sheetsToSample: Record<string, string[]>,
  sampler: (cell: AnyWorkbook) => Sample = sampleCell,
): Snapshot {
  const sheetNames = wb.worksheets.map((ws: AnyWorkbook) => ws.name as string);
  const sheets: Snapshot["sheets"] = {};
  for (const [name, addresses] of Object.entries(sheetsToSample)) {
    const ws = wb.getWorksheet(name);
    const cells: Record<string, Sample> = {};
    for (const addr of addresses) cells[addr] = sampler(ws.getCell(addr));
    const merges = [...(ws.model.merges as string[])].sort();
    sheets[name] = { cells, merges };
  }
  return { sheetNames, sheets };
}

// --- engine loading --------------------------------------------------------------------------------

async function loadRealExceljs(): Promise<{ Workbook: WorkbookCtor } | undefined> {
  const spec = "exceljs"; // variable specifier: keeps this import un-resolved statically (see header).
  try {
    const mod = (await import(spec)) as { default?: unknown } & Record<string, unknown>;
    return (mod.default ?? mod) as { Workbook: WorkbookCtor };
  } catch {
    return undefined;
  }
}

async function loadCompatExceljs(): Promise<{ Workbook: WorkbookCtor }> {
  const mod = await import("@alosha/xlsx/compat");
  return (mod.default ?? mod) as unknown as { Workbook: WorkbookCtor };
}

// --- the corpus runner -----------------------------------------------------------------------------

describe("compat parity — @alosha/xlsx/compat vs real exceljs", () => {
  for (const { name, build, sample } of corpus) {
    it(name, async (ctx) => {
      const RealExcelJS = await loadRealExceljs();
      if (!RealExcelJS) {
        ctx.skip();
        return;
      }
      const CompatExcelJS = await loadCompatExceljs();

      const real = new RealExcelJS.Workbook();
      build(real);
      const bufReal = (await real.xlsx.writeBuffer()) as Uint8Array;

      const compat = new CompatExcelJS.Workbook();
      build(compat);
      const bufCompat = (await compat.xlsx.writeBuffer()) as Uint8Array;

      // Step 1: load BOTH buffers with real exceljs and diff the sampled snapshots.
      const loadedReal = new RealExcelJS.Workbook();
      await loadedReal.xlsx.load(Buffer.from(bufReal));
      const loadedCompat = new RealExcelJS.Workbook();
      await loadedCompat.xlsx.load(Buffer.from(bufCompat));

      const referenceSnapshot = snapshot(loadedReal, sample);
      expect(snapshot(loadedCompat, sample)).toEqual(referenceSnapshot);

      // Step 2: load the compat-produced buffer with our OWN reader; same expectations.
      const ours = readWorkbookBuffer(bufCompat);
      expect(snapshot(ours, sample, sampleCoreCell)).toEqual(referenceSnapshot);
    });
  }
});
