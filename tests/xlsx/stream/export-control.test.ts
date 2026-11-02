import { getEventListeners } from "node:events";
import { existsSync } from "node:fs";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Writable } from "node:stream";
import { describe, expect, it, vi } from "vitest";
import ExcelJS from "../../../src/compat/index.js";
import {
  ExportAbortedError,
  readWorkbookBuffer,
  StreamBufferLimitError,
  WorkbookWriter,
} from "../../../src/index.js";
import { QueueSink } from "../../../src/xlsx/stream/output.js";

const turn = (): Promise<void> => new Promise((resolve) => setImmediate(resolve));

describe("stream queue byte accounting", () => {
  it("counts handed-off chunks until the iterator asks for another", async () => {
    const queue = new QueueSink(4, 10);
    const iterator = queue[Symbol.asyncIterator]();
    queue.write(new Uint8Array(4));
    let ready = false;
    const waiting = queue.whenReady().then(() => {
      ready = true;
    });
    expect((await iterator.next()).value).toHaveLength(4);
    expect(queue.bufferedBytes).toBe(4);
    await turn();
    expect(ready).toBe(false);
    const next = iterator.next();
    await waiting;
    expect(queue.bufferedBytes).toBe(0);
    queue.close();
    await next;
  });

  it("enforces the hard byte ceiling and rejects capacity waiters without retaining queued bytes", async () => {
    const queue = new QueueSink(4, 8);
    queue.write(new Uint8Array(5));
    const waiting = expect(queue.whenReady()).rejects.toBeInstanceOf(StreamBufferLimitError);
    expect(() => queue.write(new Uint8Array(4))).toThrow(StreamBufferLimitError);
    await waiting;
    expect(queue.bufferedBytes).toBe(0);
  });

  it("rejects second consumers and concurrent next calls", async () => {
    const queue = new QueueSink();
    const iterator = queue[Symbol.asyncIterator]();
    expect(() => queue[Symbol.asyncIterator]()).toThrow(/one consumer/);
    const first = iterator.next();
    await expect(iterator.next()).rejects.toThrow(/concurrent/);
    queue.close();
    await first;
  });
});

describe("export cancellation", () => {
  it("can cancel stalled finalization after all queued bytes were accepted", async () => {
    const controller = new AbortController();
    const writable = new Writable({
      write(_chunk, _encoding, callback) {
        callback();
      },
      final() {
        /* stalled finish */
      },
    });
    const writer = new WorkbookWriter({ stream: writable, signal: controller.signal });
    writer.addWorksheet("S").addRow([1]).commit();
    const committing = writer.commit();
    const failed = expect(committing).rejects.toBeInstanceOf(ExportAbortedError);
    await turn();
    controller.abort();
    await failed;
    expect(writable.destroyed).toBe(true);
    expect(getEventListeners(controller.signal, "abort")).toHaveLength(0);
  });
  it("does not open or truncate a filename when the signal was already aborted", async () => {
    const temp = await mkdtemp(join(tmpdir(), "alosha-abort-"));
    try {
      const signal = AbortSignal.abort("stop");
      const path = join(temp, "never-created.xlsx");
      const writer = new WorkbookWriter({ filename: path, signal });
      await expect(writer.commit()).rejects.toBeInstanceOf(ExportAbortedError);
      expect(existsSync(path)).toBe(false);
      expect(getEventListeners(signal, "abort")).toHaveLength(0);
    } finally {
      await rm(temp, { recursive: true, force: true });
    }
  });

  it("aborts a pending destination write, flush, and producer, and removes listeners", async () => {
    const controller = new AbortController();
    const writable = new Writable({
      highWaterMark: 1,
      write() {
        /* deliberately stalled */
      },
    });
    const writer = new WorkbookWriter({
      stream: writable,
      signal: controller.signal,
      highWaterMarkBytes: 1,
    });
    const row = writer.addWorksheet("S").addRow(["x".repeat(100_000)]);
    const writing = row.commitAsync();
    const stopped = expect(writing).rejects.toMatchObject({
      code: "EXPORT_ABORTED",
      reason: "cancel",
    });
    await turn();
    const flushing = expect(writer.flush()).rejects.toBeInstanceOf(ExportAbortedError);
    controller.abort("cancel");
    await stopped;
    await flushing;
    await expect(writer.commit()).rejects.toBeInstanceOf(ExportAbortedError);
    expect(writable.destroyed).toBe(true);
    expect(writer.bufferedBytes).toBe(0);
    await turn();
    expect(writable.listenerCount("drain")).toBe(0);
    expect(writable.listenerCount("error")).toBe(0);
    expect(getEventListeners(controller.signal, "abort")).toHaveLength(0);
  });

  it("interrupts a pending input next and requests source cleanup", async () => {
    const controller = new AbortController();
    const writer = new WorkbookWriter({ signal: controller.signal });
    const returned = vi.fn(async () => ({ done: true as const, value: undefined }));
    const rows = {
      [Symbol.asyncIterator]: () => ({
        next: () => new Promise<IteratorResult<number[]>>(() => undefined),
        return: returned,
      }),
    };
    const writing = writer.addWorksheet("S").writeRows(rows);
    const stopped = expect(writing).rejects.toBeInstanceOf(ExportAbortedError);
    await turn();
    controller.abort();
    await stopped;
    expect(returned).toHaveBeenCalledTimes(1);
  });

  it("cancels production when an async output consumer breaks early", async () => {
    const writer = new WorkbookWriter({ highWaterMarkBytes: 1 });
    const ws = writer.addWorksheet("S");
    let closed = false;
    async function* rows() {
      try {
        for (let i = 0; i < 100; i++) yield [i, "x".repeat(5000)];
      } finally {
        closed = true;
      }
    }
    const writing = ws.writeRows(rows());
    const stopped = expect(writing).rejects.toBeInstanceOf(ExportAbortedError);
    for await (const _chunk of writer) break;
    await stopped;
    await turn();
    expect(closed).toBe(true);
    expect(writer.bufferedBytes).toBe(0);
    await expect(writer.commit()).rejects.toBeInstanceOf(ExportAbortedError);
  });

  it("cancels output readers if the input source fails", async () => {
    const writer = new WorkbookWriter();
    const error = new Error("database cursor failed");
    async function* rows() {
      yield [1];
      throw error;
    }
    const consuming = (async () => {
      for await (const _chunk of writer) {
        /* drain */
      }
    })();
    const failed = expect(consuming).rejects.toBeInstanceOf(ExportAbortedError);
    await expect(writer.addWorksheet("S").writeRows(rows())).rejects.toBe(error);
    await failed;
  });

  it("removes the signal listener after successful output and ignores late aborts", async () => {
    const controller = new AbortController();
    const writer = new WorkbookWriter({
      signal: controller.signal,
      stream: new Writable({
        write(_chunk, _encoding, cb) {
          cb();
        },
      }),
    });
    await writer.addWorksheet("S").addRowAsync([1]);
    await writer.commit();
    expect(getEventListeners(controller.signal, "abort")).toHaveLength(0);
    controller.abort();
    await expect(writer.flush()).resolves.toBeUndefined();
  });
});

describe("awaitable row writes", () => {
  it("waits for a slow async-iterable consumer, keeps compressed bytes bounded, and writes a valid file", async () => {
    const writer = new WorkbookWriter({ highWaterMarkBytes: 1, maxBufferedBytes: 128 * 1024 });
    const ws = writer.addWorksheet("S");
    let produced = 0;
    let peak = 0;
    const producer = (async () => {
      for (let i = 0; i < 100; i++) {
        await ws.addRowAsync([i, "payload".repeat(1000)]);
        produced++;
        peak = Math.max(peak, writer.bufferedBytes);
      }
      await writer.commit();
    })();
    await turn();
    expect(produced).toBeLessThan(100);
    const chunks: Uint8Array[] = [];
    for await (const chunk of writer) {
      chunks.push(chunk);
      peak = Math.max(peak, writer.bufferedBytes);
      await turn();
    }
    await producer;
    expect(peak).toBeLessThanOrEqual(128 * 1024);
    expect(writer.bufferedBytes).toBe(0);
    const parsed = readWorkbookBuffer(Buffer.concat(chunks));
    expect(parsed.worksheets[0]?.rowCount).toBe(100);
    expect(parsed.worksheets[0]?.getCell("A100").value).toBe(99);
  });

  it("fails explicitly if a synchronous producer exceeds the configured buffer", () => {
    const writer = new WorkbookWriter({ highWaterMarkBytes: 1, maxBufferedBytes: 8 });
    const ws = writer.addWorksheet("S");
    expect(() => {
      ws.addRow(["x".repeat(100_000)]).commit();
      ws.commit();
    }).toThrow(StreamBufferLimitError);
    expect(writer.bufferedBytes).toBe(0);
  });

  it("validates byte-limit options", () => {
    for (const options of [
      { highWaterMarkBytes: 0 },
      { maxBufferedBytes: -1 },
      { highWaterMarkBytes: 10, maxBufferedBytes: 9 },
    ]) {
      expect(() => new WorkbookWriter(options)).toThrow(RangeError);
    }
  });

  it("allows a smaller hard ceiling without an explicit high-water mark", () => {
    const writer = new WorkbookWriter({ maxBufferedBytes: 64 * 1024 });
    expect(writer.bufferedBytes).toBe(0);
    writer.abort();
  });
});

describe("compat export control", () => {
  it("forwards awaitable rows, limits, and cancellation", async () => {
    const writer = new ExcelJS.stream.xlsx.WorkbookWriter({ maxBufferedBytes: 65536 });
    const sheet = writer.addWorksheet("S");
    await sheet.addRow([1]).commitAsync();
    await sheet.addRowAsync([2]);
    await sheet.writeRows([[3]]);
    sheet.commit();
    expect(writer.bufferedBytes).toBeGreaterThan(0);
    writer.abort("stop");
    await expect(writer.ready()).rejects.toMatchObject({ code: "EXPORT_ABORTED", reason: "stop" });
    await expect(writer.commit()).rejects.toBeInstanceOf(ExportAbortedError);
  });
});
