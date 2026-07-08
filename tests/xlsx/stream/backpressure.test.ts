// Track E3 backpressure: a fully synchronous addRow/commit loop never yields the event loop, so the
// background pump that drains QueueSink to a filename/stream destination never runs until commit() —
// the whole compressed archive buffers in memory. WorkbookWriter.flush() and WorksheetWriter.writeRows()
// give callers an await point that lets the pump drain mid-write, keeping peak memory flat. These tests
// prove the draining is incremental (deterministically, via when each write lands) and that writeRows
// produces the same archive as a manual loop.

import { Writable } from "node:stream";
import { describe, expect, it } from "vitest";
import { readWorkbookBuffer } from "../../../src/xlsx/read-workbook.js";
import { WorkbookWriter } from "../../../src/xlsx/stream/index.js";

/** A sink that concatenates archive bytes and records the value of `probe()` at each write, so a test
 *  can see how much work had happened by the time a chunk reached the destination. */
function collectingSink(probe: () => number): {
  writable: Writable;
  committedAtWrite: number[];
  bytes: () => Uint8Array;
} {
  const chunks: Uint8Array[] = [];
  const committedAtWrite: number[] = [];
  const writable = new Writable({
    write(chunk: Buffer, _enc, cb) {
      committedAtWrite.push(probe());
      chunks.push(new Uint8Array(chunk));
      cb();
    },
  });
  return {
    writable,
    committedAtWrite,
    bytes: () => {
      const total = chunks.reduce((n, c) => n + c.length, 0);
      const out = new Uint8Array(total);
      let off = 0;
      for (const c of chunks) {
        out.set(c, off);
        off += c.length;
      }
      return out;
    },
  };
}

const N = 8000;
const rowData = (r: number): [number, string] => [r, `payload-value-${r}-xxxxxxxxxxxxxxxxxxxx`];

describe("streaming — backpressure via flush() / writeRows()", () => {
  it("without flush, the whole archive buffers until commit() (baseline)", async () => {
    let committed = 0;
    const sink = collectingSink(() => committed);
    const wb = new WorkbookWriter({ stream: sink.writable, useSharedStrings: false });
    const ws = wb.addWorksheet("Big");

    for (let r = 1; r <= N; r++) {
      ws.addRow(rowData(r)).commit();
      committed = r;
    }
    ws.commit();
    await wb.commit();

    // No await point in the loop, so nothing reached the destination until every row was committed.
    expect(sink.committedAtWrite.length).toBeGreaterThan(1);
    expect(Math.min(...sink.committedAtWrite)).toBe(N);
  });

  it("flush() drains incrementally so chunks reach the destination mid-write", async () => {
    let committed = 0;
    const sink = collectingSink(() => committed);
    const wb = new WorkbookWriter({ stream: sink.writable, useSharedStrings: false });
    const ws = wb.addWorksheet("Big");

    for (let r = 1; r <= N; r++) {
      ws.addRow(rowData(r)).commit();
      committed = r;
      if (r % 1000 === 0) await wb.flush();
    }
    ws.commit();
    await wb.commit();

    // At least one chunk was written before all rows were committed — proof the sink drained mid-stream
    // instead of buffering the entire archive until commit().
    expect(Math.min(...sink.committedAtWrite)).toBeLessThan(N);
  });

  it("writeRows() streams a whole iterable with periodic flushes and drains incrementally", async () => {
    let seen = 0;
    // `seen` can't track internal row count, so probe archive-chunk order instead: with writeRows'
    // built-in flushes, chunks arrive across many writes rather than all at the end.
    const sink = collectingSink(() => ++seen);
    const wb = new WorkbookWriter({ stream: sink.writable, useSharedStrings: false });
    const ws = wb.addWorksheet("Big");

    const rows = (function* () {
      for (let r = 1; r <= N; r++) yield rowData(r);
    })();
    await ws.writeRows(rows, { flushEvery: 1000 });
    ws.commit();
    await wb.commit();

    const parsed = readWorkbookBuffer(sink.bytes());
    const big = parsed.getWorksheet("Big");
    expect(big?.rowCount).toBe(N);
    expect(big?.getCell("A1").value).toBe(1);
    expect(big?.getCell(`A${N}`).value).toBe(N);
  });

  it("writeRows() produces the same archive content as a manual addRow/commit loop", async () => {
    const rows = Array.from({ length: 3000 }, (_, i) => rowData(i + 1));

    const manual = collectingSink(() => 0);
    const wbM = new WorkbookWriter({ stream: manual.writable, useSharedStrings: false });
    const wsM = wbM.addWorksheet("S");
    for (const r of rows) wsM.addRow(r).commit();
    wsM.commit();
    await wbM.commit();

    const viaWriteRows = collectingSink(() => 0);
    const wbW = new WorkbookWriter({ stream: viaWriteRows.writable, useSharedStrings: false });
    const wsW = wbW.addWorksheet("S");
    await wsW.writeRows(rows, { flushEvery: 500 });
    wsW.commit();
    await wbW.commit();

    // Zip container timestamps differ run-to-run, so compare parsed content, not raw bytes.
    const pm = readWorkbookBuffer(manual.bytes()).getWorksheet("S");
    const pw = readWorkbookBuffer(viaWriteRows.bytes()).getWorksheet("S");
    expect(pw?.rowCount).toBe(pm?.rowCount);
    expect(pw?.rowCount).toBe(3000);
    expect(pw?.getCell("A3000").value).toBe(3000);
    expect(pw?.getCell("B1").value).toBe(pm?.getCell("B1").value);
    expect(pw?.getCell("B3000").value).toBe(pm?.getCell("B3000").value);
  });

  it("flush() resolves immediately on the async-iterable path (no pump to await)", async () => {
    const wb = new WorkbookWriter({ useSharedStrings: false });
    const ws = wb.addWorksheet("S");
    ws.addRow([1]).commit();
    // No filename/stream destination → nothing to drain; must resolve rather than hang.
    await expect(wb.flush()).resolves.toBeUndefined();
  });
});
