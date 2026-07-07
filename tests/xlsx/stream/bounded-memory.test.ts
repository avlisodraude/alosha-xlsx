// Track E2 bounded-memory smoke: the streaming writer must not retain committed rows, so a large sheet
// serialises with roughly flat memory (the whole point of streaming — see bench/streaming.ts for the
// throughput/RSS numbers). We prove non-retention deterministically rather than with a flaky heap
// threshold: hold only a WeakRef to an early committed row, stream tens of thousands more, force a full
// GC, and assert the WeakRef has been collected. If the writer held onto committed rows, the first row
// would stay reachable and the deref would be non-undefined.

import { Writable } from "node:stream";
import v8 from "node:v8";
import vm from "node:vm";
import { describe, expect, it } from "vitest";
import { readWorkbookBuffer } from "../../../src/xlsx/read-workbook.js";
import { WorkbookWriter } from "../../../src/xlsx/stream/index.js";

// Expose a real GC without the `--expose-gc` CLI flag, so this test is self-contained under `pnpm test`.
v8.setFlagsFromString("--expose-gc");
const gc = vm.runInNewContext("gc") as () => void;

const ROWS = 20_000;

/** A sink that concatenates the archive bytes (compressed output is small) but retains no rows. */
function collectingSink(): { writable: Writable; bytes: () => Uint8Array } {
  const chunks: Uint8Array[] = [];
  const writable = new Writable({
    write(chunk: Buffer, _enc, cb) {
      chunks.push(new Uint8Array(chunk));
      cb();
    },
  });
  return {
    writable,
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

/** Commit one row and return only a WeakRef to it — the strong reference dies with this call frame. */
function commitTracked(ws: ReturnType<WorkbookWriter["addWorksheet"]>): WeakRef<object> {
  const row = ws.addRow([1, "first", true]);
  row.commit();
  return new WeakRef(row as unknown as object);
}

describe("streaming — bounded memory", () => {
  it("does not retain committed rows across a large sheet", async () => {
    const sink = collectingSink();
    const wb = new WorkbookWriter({ stream: sink.writable, useSharedStrings: false });
    const ws = wb.addWorksheet("Big");

    const firstRef = commitTracked(ws);
    for (let r = 2; r <= ROWS; r++) {
      ws.addRow([r, `row-${r}`, r % 2 === 0]).commit();
    }
    ws.commit();
    await wb.commit();

    // Force a full GC (twice, across a turn of the loop) and confirm the early row was collectable —
    // it cannot be if the writer still holds it.
    gc();
    await new Promise((resolve) => setImmediate(resolve));
    gc();
    expect(firstRef.deref()).toBeUndefined();

    // And every streamed row actually made it into a valid archive (nothing silently dropped).
    const parsed = readWorkbookBuffer(sink.bytes());
    const big = parsed.getWorksheet("Big");
    expect(big?.rowCount).toBe(ROWS);
    expect(big?.getCell("A1").value).toBe(1);
    expect(big?.getCell(`A${ROWS}`).value).toBe(ROWS);
  });
});
