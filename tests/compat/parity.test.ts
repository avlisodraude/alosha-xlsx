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
  {
    name: "hyperlink cells (external URL, mailto, and a tooltip)",
    build(wb) {
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = { text: "Anthropic", hyperlink: "https://anthropic.com" };
      ws.getCell("A2").value = {
        text: "Mail us",
        hyperlink: "mailto:hi@example.com",
        tooltip: "Send an email",
      };
      ws.getCell("A3").value = "plain text";
    },
    sample: { Sheet1: ["A1", "A2", "A3"] },
  },
  {
    name: "cell.dataValidation list dropdown",
    build(wb) {
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = "pick one";
      ws.getCell("A1").dataValidation = {
        type: "list",
        allowBlank: true,
        formulae: ['"Yes,No,Maybe"'],
      };
      ws.getCell("A2").value = "no validation";
    },
    sample: { Sheet1: ["A1", "A2"] },
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
  dataValidation: unknown;
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
    if ("text" in obj && "hyperlink" in obj) {
      // ExcelJS's own reader drops `tooltip` when resolving a hyperlink cell (it keeps only the
      // rId's resolved target, see `hyperlinkMap` in its worksheet-xform) even though it does
      // write the attribute — so the Step 1 reference (round-tripped through real ExcelJS) can
      // never carry one. Compare only { text, hyperlink } here; tooltip fidelity for OUR reader
      // is asserted directly in read-foreign.test.ts / read-roundtrip.test.ts instead.
      const h = obj as { text: string; hyperlink: string };
      return { text: h.text, hyperlink: h.hyperlink };
    }
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
    dataValidation: cell.dataValidation ?? null,
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

// --- panes + auto-filter (PANES-AUTOFILTER-SPEC.md §4) -------------------------------------------
//
// `views`/`autoFilter` are sheet-level, not per-cell, so they sit outside the cell-sampling corpus
// above; same two-engine build-then-diff shape, sampling the sheet view/filter instead of cells.

interface ViewSample {
  state: string | undefined;
  xSplit: number | undefined;
  ySplit: number | undefined;
  autoFilter: string | null;
}

function sampleView(wb: AnyWorkbook): ViewSample {
  const ws = wb.getWorksheet("Sheet1");
  const view = (ws.views[0] ?? {}) as { state?: string; xSplit?: number; ySplit?: number };
  return {
    state: view.state,
    xSplit: view.xSplit,
    ySplit: view.ySplit,
    autoFilter: ws.autoFilter ?? null,
  };
}

describe("compat parity — frozen pane + auto-filter", () => {
  it("views[0].state/xSplit/ySplit and the autoFilter range match real exceljs", async (ctx) => {
    const RealExcelJS = await loadRealExceljs();
    if (!RealExcelJS) {
      ctx.skip();
      return;
    }
    const CompatExcelJS = await loadCompatExceljs();

    const build: Build = (wb) => {
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = "header";
      ws.views = [{ state: "frozen", xSplit: 1, ySplit: 1, topLeftCell: "B2" }];
      ws.autoFilter = "A1:D1";
    };

    const real = new RealExcelJS.Workbook();
    build(real);
    const bufReal = (await real.xlsx.writeBuffer()) as Uint8Array;

    const compat = new CompatExcelJS.Workbook();
    build(compat);
    const bufCompat = (await compat.xlsx.writeBuffer()) as Uint8Array;

    const loadedReal = new RealExcelJS.Workbook();
    await loadedReal.xlsx.load(Buffer.from(bufReal));
    const loadedCompat = new RealExcelJS.Workbook();
    await loadedCompat.xlsx.load(Buffer.from(bufCompat));

    expect(sampleView(loadedCompat)).toEqual(sampleView(loadedReal));
  });

  it("normalizes the {from,to} object form of autoFilter to an A1 range string", async () => {
    const CompatExcelJS = await loadCompatExceljs();
    const wb = new CompatExcelJS.Workbook();

    const bySheetAddress = wb.addWorksheet("BySheetAddress");
    bySheetAddress.autoFilter = { from: "A1", to: "D1" };
    expect(bySheetAddress.autoFilter).toBe("A1:D1");

    const byRowColumn = wb.addWorksheet("ByRowColumn");
    byRowColumn.autoFilter = { from: { row: 1, column: 1 }, to: { row: 3, column: 4 } };
    expect(byRowColumn.autoFilter).toBe("A1:D3");
  });
});

// --- conditional formatting (CONDITIONAL-FORMATTING-SPEC.md §4) ----------------------------------
//
// Sheet-level like views/autoFilter, so it sits outside the cell-sampling corpus too. A cellIs
// (highlight fill, the dxf-bearing family) + a colorScale (the {argb}-colour self-visualising
// family) — same two-engine build-then-diff shape, diffed on the fields that survive an exceljs
// round trip: rule type/operator/priority, the highlight fill's resolved argb, and the colorScale's
// resolved colours. `formulae`/`cfvo` are left out of the diff — real exceljs's own reader returns a
// cellIs formula as a string (`"5"`), not the number our own reader coerces it to, an orthogonal
// stringly-typed-XML quirk that has nothing to do with the compat shim itself.

interface CfSample {
  ref: string;
  type: string;
  operator: string | undefined;
  priority: number | undefined;
  fillArgb: string | undefined;
  colors: (string | undefined)[] | undefined;
}

function sampleConditionalFormattings(wb: AnyWorkbook, sheetName: string): CfSample[] {
  const ws = wb.getWorksheet(sheetName);
  const cfs = ws.conditionalFormattings as Array<{
    ref: string;
    rules: Array<{
      type: string;
      operator?: string;
      priority?: number;
      style?: { fill?: { fgColor?: { argb?: string }; bgColor?: { argb?: string } } };
      color?: Array<{ argb?: string }>;
    }>;
  }>;
  const samples: CfSample[] = [];
  for (const cf of cfs) {
    for (const rule of cf.rules) {
      const fill = rule.style?.fill;
      samples.push({
        ref: cf.ref,
        type: rule.type,
        operator: rule.operator,
        priority: rule.priority,
        fillArgb: fill?.bgColor?.argb ?? fill?.fgColor?.argb,
        colors: rule.color?.map((c) => c?.argb),
      });
    }
  }
  return samples;
}

// --- cell comments (COMMENTS-SPEC.md §3) ----------------------------------------------------------
//
// Notes sit outside the cell-sampling corpus (like views/CF) because real exceljs's own reader
// normalizes them in ways the generic Sample shape doesn't capture: its comments-xform hardcodes a
// single "Author" on write and never resolves one back onto `note.author` on read, and a single
// plain-text run collapses back into a bare string while a multi-run note stays an object carrying
// its runs array plus exceljs's own margins/protection/editAs boilerplate. So the diff here is
// deliberately narrow — the concatenated run text of each noted cell. Per the C1 finding, exceljs
// only reads a note on a cell that also holds a value, so every noted cell below is valued too.

function sampleNoteText(note: unknown): string | undefined {
  if (note === undefined) return undefined;
  if (typeof note === "string") return note;
  const n = note as { texts?: Array<{ text: string }> };
  return (n.texts ?? []).map((t) => t.text).join("");
}

function sampleNotes(
  wb: AnyWorkbook,
  sheetName: string,
  addresses: string[],
): Record<string, string | undefined> {
  const ws = wb.getWorksheet(sheetName);
  const out: Record<string, string | undefined> = {};
  for (const addr of addresses) out[addr] = sampleNoteText(ws.getCell(addr).note);
  return out;
}

describe("compat parity — cell comments", () => {
  it("a plain-string note and a rich-text note match real exceljs after reload", async (ctx) => {
    const RealExcelJS = await loadRealExceljs();
    if (!RealExcelJS) {
      ctx.skip();
      return;
    }
    const CompatExcelJS = await loadCompatExceljs();

    const build: Build = (wb) => {
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = "annotated";
      ws.getCell("A1").note = "hello";
      ws.getCell("B1").value = "styled";
      ws.getCell("B1").note = {
        texts: [{ text: "bold", font: { bold: true } }, { text: " plain" }],
      };
    };

    const real = new RealExcelJS.Workbook();
    build(real);
    const bufReal = (await real.xlsx.writeBuffer()) as Uint8Array;

    const compat = new CompatExcelJS.Workbook();
    build(compat);
    const bufCompat = (await compat.xlsx.writeBuffer()) as Uint8Array;

    const loadedReal = new RealExcelJS.Workbook();
    await loadedReal.xlsx.load(Buffer.from(bufReal));
    const loadedCompat = new RealExcelJS.Workbook();
    await loadedCompat.xlsx.load(Buffer.from(bufCompat));

    expect(sampleNotes(loadedCompat, "Sheet1", ["A1", "B1"])).toEqual(
      sampleNotes(loadedReal, "Sheet1", ["A1", "B1"]),
    );

    // Our own reader recovers both too, with the author defaulted (exceljs never round-trips one).
    const ours = readWorkbookBuffer(bufCompat);
    const oursWs = ours.getWorksheet("Sheet1");
    expect(oursWs?.getCell("A1").note).toEqual({
      texts: { richText: [{ text: "hello" }] },
      author: "Author",
    });
    expect(oursWs?.getCell("B1").note).toEqual({
      texts: { richText: [{ text: "bold", font: { bold: true } }, { text: " plain" }] },
      author: "Author",
    });
  });
});

describe("compat parity — conditional formatting (cellIs highlight + colorScale)", () => {
  it("rule type/operator/priority, the highlight argb, and the colorScale colours match real exceljs", async (ctx) => {
    const RealExcelJS = await loadRealExceljs();
    if (!RealExcelJS) {
      ctx.skip();
      return;
    }
    const CompatExcelJS = await loadCompatExceljs();

    const build: Build = (wb) => {
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = 10;
      ws.addConditionalFormatting({
        ref: "A1:A10",
        rules: [
          {
            type: "cellIs",
            operator: "greaterThan",
            formulae: [5],
            priority: 1,
            style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFF8696B" } } },
          },
        ],
      });
      ws.addConditionalFormatting({
        ref: "C1:C20",
        rules: [
          {
            type: "colorScale",
            priority: 2,
            cfvo: [{ type: "min" }, { type: "max" }],
            color: [{ argb: "FFF8696B" }, { argb: "FF63BE7B" }],
          },
        ],
      });
    };

    const real = new RealExcelJS.Workbook();
    build(real);
    const bufReal = (await real.xlsx.writeBuffer()) as Uint8Array;

    const compat = new CompatExcelJS.Workbook();
    build(compat);
    const bufCompat = (await compat.xlsx.writeBuffer()) as Uint8Array;

    const loadedReal = new RealExcelJS.Workbook();
    await loadedReal.xlsx.load(Buffer.from(bufReal));
    const loadedCompat = new RealExcelJS.Workbook();
    await loadedCompat.xlsx.load(Buffer.from(bufCompat));

    expect(sampleConditionalFormattings(loadedCompat, "Sheet1")).toEqual(
      sampleConditionalFormattings(loadedReal, "Sheet1"),
    );
  });
});
