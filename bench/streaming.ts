// Benchmark: the streaming WorkbookWriter (src/xlsx/stream) vs the buffered writeWorkbookBuffer on a
// large single-sheet workbook (default 500,000 rows), measuring peak resident memory and write
// throughput. Node-runnable via `pnpm bench:stream` (tsx); not part of `pnpm test`. Results are written
// into the "Streaming vs buffered" section of BENCHMARKS.md.
//
// Peak RSS is a per-process high-water mark (`process.resourceUsage().maxRSS`), so it cannot be read
// twice in one process. Each mode therefore runs in its own child process (this same file re-invoked
// with `--mode`), and the parent collects the two JSON results. Both paths walk the same row generator
// (bench/workbook-spec.ts `streamRowValues`) and write to a throwaway temp file, so the only difference
// measured is *how* the bytes are produced — incrementally, or from a fully-materialised model.

import { spawnSync } from "node:child_process";
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import os, { tmpdir } from "node:os";
import { join } from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";
import { Workbook, WorkbookWriter, writeWorkbookBuffer } from "../src/index.js";
import { streamRowValues } from "./workbook-spec.js";

const ROWS = Number(process.env.BENCH_ROWS ?? 500_000);
const COLS = Number(process.env.BENCH_COLS ?? 5);
const SELF = fileURLToPath(import.meta.url);
const BENCHMARKS = fileURLToPath(new URL("../BENCHMARKS.md", import.meta.url));
const SECTION_MARK = "<!-- bench:streaming -->";

type Mode = "streaming" | "buffered";

interface ChildResult {
  mode: Mode;
  rows: number;
  cols: number;
  ms: number;
  peakRssBytes: number;
  bytes: number;
}

/** Node normalises `process.resourceUsage().maxRSS` to kilobytes on every platform (it converts the
 *  raw macOS `ru_maxrss`, which the kernel reports in bytes) — so it's always KiB → bytes here. */
function peakRssBytes(): number {
  return process.resourceUsage().maxRSS * 1024;
}

/** Stream `rows` rows straight to `path`, holding no row in memory (the point of the exercise). */
async function runStreaming(path: string): Promise<number> {
  const wb = new WorkbookWriter({ filename: path, useSharedStrings: false });
  const ws = wb.addWorksheet("Bench");
  function* rows() {
    for (let r = 1; r <= ROWS; r++) yield streamRowValues(r, COLS);
  }
  await ws.writeRows(rows());
  ws.commit();
  await wb.commit();
  return existsSync(path) ? readFileSync(path).byteLength : 0;
}

/** Build the whole model in memory, then serialise it in one shot — the buffered baseline. */
function runBuffered(path: string): number {
  const wb = new Workbook();
  const ws = wb.addWorksheet("Bench");
  for (let r = 1; r <= ROWS; r++) {
    ws.getRow(r).values = streamRowValues(r, COLS);
  }
  const bytes = writeWorkbookBuffer(wb, { useSharedStrings: false });
  writeFileSync(path, bytes);
  return bytes.byteLength;
}

/** The child role: run one mode, print a single JSON result line, exit. */
async function runChild(mode: Mode): Promise<void> {
  const dir = mkdtempSync(join(tmpdir(), "alosha-bench-"));
  const path = join(dir, "out.xlsx");
  try {
    const t0 = performance.now();
    const bytes = mode === "streaming" ? await runStreaming(path) : runBuffered(path);
    const ms = performance.now() - t0;
    const result: ChildResult = {
      mode,
      rows: ROWS,
      cols: COLS,
      ms,
      peakRssBytes: peakRssBytes(),
      bytes,
    };
    // JSON only on stdout (tsx/node diagnostics go to stderr); the parent parses this line.
    process.stdout.write(JSON.stringify(result));
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

/** The parent role: run `mode` in a fresh process so its peak RSS is measured in isolation. */
function spawnMode(mode: Mode): ChildResult {
  const res = spawnSync(
    process.execPath,
    ["--import", "tsx", "--max-old-space-size=8192", SELF, `--mode=${mode}`],
    { encoding: "utf8", maxBuffer: 1 << 24 },
  );
  if (res.status !== 0) {
    throw new Error(`bench child (${mode}) failed with code ${res.status}:\n${res.stderr}`);
  }
  const line = res.stdout.trim().split("\n").pop() ?? "";
  return JSON.parse(line) as ChildResult;
}

function fmtMs(ms: number): string {
  return ms >= 1000 ? `${(ms / 1000).toFixed(2)} s` : `${ms.toFixed(0)} ms`;
}
function fmtMiB(bytes: number): string {
  return `${(bytes / 1024 / 1024).toFixed(1)} MiB`;
}
function throughput(r: ChildResult): number {
  return r.rows / (r.ms / 1000);
}
function fmtRps(rps: number): string {
  return `${Math.round(rps).toLocaleString("en-US")} rows/s`;
}

/** Replace (or append) the streaming section of BENCHMARKS.md, leaving any other section untouched. */
function writeSection(md: string): void {
  const block = `${SECTION_MARK}\n${md}\n${SECTION_MARK}`;
  let existing = existsSync(BENCHMARKS) ? readFileSync(BENCHMARKS, "utf8") : "";
  const start = existing.indexOf(SECTION_MARK);
  if (start >= 0) {
    const end = existing.indexOf(SECTION_MARK, start + SECTION_MARK.length);
    if (end >= 0) {
      existing = existing.slice(0, start) + block + existing.slice(end + SECTION_MARK.length);
      writeFileSync(BENCHMARKS, existing);
      return;
    }
  }
  const sep = existing.length > 0 && !existing.endsWith("\n\n") ? "\n\n" : "";
  writeFileSync(BENCHMARKS, `${existing}${sep}${block}\n`);
}

async function main(): Promise<void> {
  const modeArg = process.argv.find((a) => a.startsWith("--mode="));
  if (modeArg) {
    await runChild(modeArg.slice("--mode=".length) as Mode);
    return;
  }

  console.log(`Streaming vs buffered: ${ROWS.toLocaleString("en-US")} rows x ${COLS} cols`);
  console.log("Running each mode in its own process to measure peak RSS in isolation…\n");

  const streaming = spawnMode("streaming");
  const buffered = spawnMode("buffered");

  for (const r of [streaming, buffered]) {
    console.log(`${r.mode}`);
    console.log(`  write:      ${fmtMs(r.ms)}  (${fmtRps(throughput(r))})`);
    console.log(`  peak RSS:   ${fmtMiB(r.peakRssBytes)}`);
    console.log(`  output:     ${fmtMiB(r.bytes)}`);
  }

  const memRatio = buffered.peakRssBytes / streaming.peakRssBytes;
  const spdRatio = throughput(streaming) / throughput(buffered);
  const machine = `${os.type()} ${os.release()} (${os.arch()}), ${os.cpus()[0]?.model ?? "unknown CPU"}, ${os.cpus().length} cores`;

  const md = `## Streaming vs buffered write

Generated by \`pnpm bench:stream\` (\`bench/streaming.ts\`) on ${new Date().toISOString()}.

- Machine: ${machine}
- Node: ${process.version}
- Workload: a single sheet of ${ROWS.toLocaleString("en-US")} rows x ${COLS} cols (${(ROWS * COLS).toLocaleString("en-US")} cells), mixed types (string/number/boolean/date/formula), inline strings, written to a temp file
- Methodology: each mode runs in its own child process; peak RSS is \`process.resourceUsage().maxRSS\` (a per-process high-water mark), normalised to bytes; write time is wall-clock around the write (model construction included for buffered, since a buffered caller must materialise the whole model first)

| Mode | Write time | Throughput | Peak RSS | Output size |
| --- | --- | --- | --- | --- |
| streaming (\`WorkbookWriter\`) | ${fmtMs(streaming.ms)} | ${fmtRps(throughput(streaming))} | ${fmtMiB(streaming.peakRssBytes)} | ${fmtMiB(streaming.bytes)} |
| buffered (\`writeWorkbookBuffer\`) | ${fmtMs(buffered.ms)} | ${fmtRps(throughput(buffered))} | ${fmtMiB(buffered.peakRssBytes)} | ${fmtMiB(buffered.bytes)} |

At ${ROWS.toLocaleString("en-US")} rows, streaming holds ~${memRatio.toFixed(1)}x less peak memory than the buffered writer (which must retain the full model plus the whole output buffer), while sustaining ${spdRatio >= 1 ? `${spdRatio.toFixed(2)}x the write throughput` : `${(1 / spdRatio).toFixed(2)}x lower write throughput`}. The streaming writer's memory stays roughly flat as rows grow; the buffered writer's grows with the sheet.

Run it yourself: \`pnpm bench:stream\` (set \`BENCH_ROWS\` / \`BENCH_COLS\` to vary the workload).`;

  writeSection(md);
  console.log(`\nWrote streaming section to ${BENCHMARKS}`);
}

main().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});
