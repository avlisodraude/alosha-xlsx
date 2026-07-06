// Benchmark: @alosha/xlsx vs exceljs (the existing devDependency) writing and reading back the same
// mid-size workbook (see workbook-spec.ts — 50,000 cells, mixed types/styles/merges). Node-runnable via
// `pnpm bench` (tsx); deliberately not part of `pnpm test`. Results are written to BENCHMARKS.md.
//
// Not timed: building the in-memory workbook from the spec (model construction). Timed: only the
// write (serialize to bytes) and read-back (parse bytes into a fresh workbook) steps, since that is
// what a consumer actually pays for per file.

import { writeFileSync } from "node:fs";
import os from "node:os";
import process from "node:process";
import { fileURLToPath } from "node:url";
import { Workbook } from "../src/index.js";
import { buildWorkbookSpec, type WorkbookSpec } from "./workbook-spec.js";

const ROWS = 2000;
const COLS = 25;
const RUNS = 5;
const WARMUP = 1;

function applyAlosha(wb: Workbook, spec: WorkbookSpec): void {
  const ws = wb.addWorksheet("Bench");
  for (const c of spec.cells) {
    const cell = ws.getCell(c.row, c.col);
    cell.value = c.value;
    if (c.style?.font) cell.font = c.style.font;
    if (c.style?.fill)
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.style.fill.fgArgb } };
    if (c.style?.numFmt) cell.numFmt = c.style.numFmt;
  }
  for (const m of spec.merges) ws.mergeCells(m.top, m.left, m.bottom, m.right);
}

// biome-ignore lint/suspicious/noExplicitAny: exceljs has no static types available here (see tests/xlsx/write-workbook.test.ts).
function applyExceljs(wb: any, spec: WorkbookSpec): void {
  const ws = wb.addWorksheet("Bench");
  for (const c of spec.cells) {
    const cell = ws.getCell(c.row, c.col);
    cell.value = c.value;
    if (c.style?.font) cell.font = c.style.font;
    if (c.style?.fill) {
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.style.fill.fgArgb } };
    }
    if (c.style?.numFmt) cell.numFmt = c.style.numFmt;
  }
  for (const m of spec.merges) ws.mergeCells(m.top, m.left, m.bottom, m.right);
}

interface RunResult {
  writeMs: number;
  readMs: number;
  bytes: number;
}

async function benchAlosha(spec: WorkbookSpec): Promise<RunResult[]> {
  const results: RunResult[] = [];
  for (let i = 0; i < WARMUP + RUNS; i++) {
    const wb = new Workbook();
    applyAlosha(wb, spec);

    const t0 = performance.now();
    const buf = await wb.xlsx.writeBuffer();
    const t1 = performance.now();

    const wb2 = new Workbook();
    await wb2.xlsx.load(buf);
    const t2 = performance.now();

    if (i >= WARMUP) results.push({ writeMs: t1 - t0, readMs: t2 - t1, bytes: buf.byteLength });
  }
  return results;
}

async function benchExceljs(spec: WorkbookSpec): Promise<RunResult[]> {
  const mod = (await import("exceljs")) as { default?: unknown } & Record<string, unknown>;
  // biome-ignore lint/suspicious/noExplicitAny: exceljs has no static types available here.
  const ExcelJS = (mod.default ?? mod) as any;

  const results: RunResult[] = [];
  for (let i = 0; i < WARMUP + RUNS; i++) {
    const wb = new ExcelJS.Workbook();
    applyExceljs(wb, spec);

    const t0 = performance.now();
    const buf = await wb.xlsx.writeBuffer();
    const t1 = performance.now();

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.load(buf);
    const t2 = performance.now();

    if (i >= WARMUP) results.push({ writeMs: t1 - t0, readMs: t2 - t1, bytes: buf.byteLength });
  }
  return results;
}

function median(values: number[]): number {
  const sorted = [...values].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  const m =
    sorted.length % 2 === 0
      ? ((sorted[mid - 1] as number) + (sorted[mid] as number)) / 2
      : (sorted[mid] as number);
  return m;
}

function summarize(name: string, results: RunResult[]) {
  return {
    name,
    writeMsMedian: median(results.map((r) => r.writeMs)),
    readMsMedian: median(results.map((r) => r.readMs)),
    bytes: results[0]?.bytes ?? 0,
    runs: results.length,
  };
}

function fmtMs(ms: number): string {
  return `${ms.toFixed(1)} ms`;
}

function fmtBytes(bytes: number): string {
  return `${(bytes / 1024).toFixed(1)} KiB`;
}

async function main() {
  const spec = buildWorkbookSpec(ROWS, COLS);
  console.log(
    `Workbook spec: ${spec.rows} rows x ${spec.cols} cols = ${spec.cellCount} cells, ${spec.merges.length} merges`,
  );
  console.log(`Runs per library: ${RUNS} (+${WARMUP} warmup, discarded)\n`);

  const aloshaResults = await benchAlosha(spec);
  const exceljsResults = await benchExceljs(spec);

  const alosha = summarize("@alosha/xlsx", aloshaResults);
  const exceljs = summarize("exceljs", exceljsResults);

  for (const s of [alosha, exceljs]) {
    console.log(`${s.name}`);
    console.log(`  write (median of ${s.runs}): ${fmtMs(s.writeMsMedian)}`);
    console.log(`  read  (median of ${s.runs}): ${fmtMs(s.readMsMedian)}`);
    console.log(`  output size: ${fmtBytes(s.bytes)}`);
  }

  const machine = `${os.type()} ${os.release()} (${os.arch()}), ${os.cpus()[0]?.model ?? "unknown CPU"}, ${os.cpus().length} cores`;
  const generatedAt = new Date().toISOString();

  // aMs/aBytes is @alosha/xlsx, bMs/bBytes is exceljs; phrase relative to alosha so "faster"/"larger"
  // always means "alosha is faster/larger than exceljs" and can't be misread like a bare ratio can.
  const compareTime = (aMs: number, bMs: number) => {
    const ratio = aMs <= bMs ? bMs / aMs : aMs / bMs;
    return aMs <= bMs ? `${ratio.toFixed(2)}x faster` : `${ratio.toFixed(2)}x slower`;
  };
  const compareSize = (aBytes: number, bBytes: number) => {
    const ratio = aBytes >= bBytes ? aBytes / bBytes : bBytes / aBytes;
    return aBytes >= bBytes ? `${ratio.toFixed(2)}x larger` : `${ratio.toFixed(2)}x smaller`;
  };

  const md = `# Benchmarks

Generated by \`pnpm bench\` (\`bench/index.ts\`) on ${generatedAt}.

- Machine: ${machine}
- Node: ${process.version}
- Workbook: ${spec.rows} rows x ${spec.cols} cols = ${spec.cellCount} cells, mixed types (string/number/boolean/date/formula), scattered bold/fill/numFmt styles, ${spec.merges.length} merges
- Methodology: ${RUNS} timed runs per library (+${WARMUP} discarded warmup run); each run builds a fresh workbook, times \`writeBuffer()\` and a fresh \`load()\` of the resulting bytes separately; medians reported

| Library | Write (median) | Read (median) | Output size |
| --- | --- | --- | --- |
| @alosha/xlsx | ${fmtMs(alosha.writeMsMedian)} | ${fmtMs(alosha.readMsMedian)} | ${fmtBytes(alosha.bytes)} |
| exceljs | ${fmtMs(exceljs.writeMsMedian)} | ${fmtMs(exceljs.readMsMedian)} | ${fmtBytes(exceljs.bytes)} |

On this workbook, @alosha/xlsx write is ${compareTime(alosha.writeMsMedian, exceljs.writeMsMedian)} than exceljs, read is ${compareTime(alosha.readMsMedian, exceljs.readMsMedian)}, and its output is ${compareSize(alosha.bytes, exceljs.bytes)}.

Run it yourself: \`pnpm bench\`.
`;

  const outPath = fileURLToPath(new URL("../BENCHMARKS.md", import.meta.url));
  writeFileSync(outPath, md);
  console.log(`\nWrote ${outPath}`);
}

main().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});
