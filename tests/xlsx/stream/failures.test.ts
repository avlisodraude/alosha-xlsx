import { Writable } from "node:stream";
import { describe, expect, it } from "vitest";
import { WorkbookWriter } from "../../../src/xlsx/stream/index.js";
import { drainToWritable, QueueSink } from "../../../src/xlsx/stream/output.js";

describe("streaming destination failures", () => {
  it("rejects flush and commit, and stops rows after a destination write fails", async () => {
    const error = new Error("disk full");
    const writable = new Writable({
      highWaterMark: 1,
      write(_chunk, _encoding, callback) {
        setImmediate(() => callback(error));
      },
    });
    const wb = new WorkbookWriter({ stream: writable });
    const ws = wb.addWorksheet("S");
    ws.addRow(["value"]).commit();
    ws.commit();
    await expect(wb.flush()).rejects.toBe(error);
    await expect(wb.commit()).rejects.toBe(error);
    expect(() => wb.addWorksheet("Other")).toThrow(error);
    expect(writable.listenerCount("drain")).toBe(0);
    expect(writable.listenerCount("error")).toBe(0);
  });

  it("detects a premature close while waiting for producer bytes", async () => {
    const writable = new Writable({
      write(_chunk, _encoding, cb) {
        cb();
      },
    });
    const wb = new WorkbookWriter({ stream: writable });
    writable.destroy();
    await new Promise((resolve) => setImmediate(resolve));
    await expect(wb.flush()).rejects.toThrow("closed before completion");
    await expect(wb.commit()).rejects.toThrow("closed before completion");
  });

  it("detects errors while idle, even when no write is backpressured", async () => {
    const writable = new Writable({
      write(_chunk, _encoding, cb) {
        cb();
      },
    });
    const wb = new WorkbookWriter({ stream: writable });
    const error = new Error("connection lost");
    writable.destroy(error);
    await new Promise((resolve) => setImmediate(resolve));
    await expect(wb.commit()).rejects.toBe(error);
  });

  it("waits for a backpressured final chunk before resolving flush", async () => {
    const source = new QueueSink();
    let release!: () => void;
    const writable = new Writable({
      highWaterMark: 1,
      write(_chunk, _encoding, cb) {
        release = cb;
      },
    });
    const pumped = drainToWritable(source, writable);
    source.write(new Uint8Array([1]));
    let drained = false;
    const flushed = source.whenDrained().then(() => {
      drained = true;
    });
    await new Promise((resolve) => setImmediate(resolve));
    expect(drained).toBe(false);
    release();
    await flushed;
    source.close();
    await pumped;
  });

  it("rejects filename errors without leaving background rejections", async () => {
    const wb = new WorkbookWriter({ filename: "/dev/null/alosha-export.xlsx" });
    wb.addWorksheet("S").addRow([1]).commit();
    await expect(wb.commit()).rejects.toThrow();
  });
});
