import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import {
  DEFAULT_READ_LIMITS,
  ReadLimitError,
  readWorkbookBuffer,
  Workbook,
  writeWorkbookBuffer,
} from "../../src/index.js";
import { readZipArchive, writeZipArchive } from "../../src/zip-layer.js";

const archive = writeZipArchive(
  new Map([
    ["first", new Uint8Array(1024)],
    ["second", new Uint8Array(1024)],
  ]),
);

describe("read budgets", () => {
  it.each([
    ["maxInputBytes", archive.length - 1],
    ["maxEntries", 1],
    ["maxEntryBytes", 1023],
    ["maxTotalBytes", 2047],
  ] as const)("rejects %s before exceeding its budget", (limit, maximum) => {
    try {
      readWorkbookBuffer(archive, { [limit]: maximum });
      throw new Error("expected read limit rejection");
    } catch (error) {
      expect(error).toBeInstanceOf(ReadLimitError);
      expect(error).toMatchObject({ code: "READ_LIMIT_EXCEEDED", limit, maximum });
    }
  });

  it("allows exact boundaries and explicit overrides", () => {
    expect(
      readZipArchive(archive, {
        maxInputBytes: archive.length,
        maxEntries: 2,
        maxEntryBytes: 1024,
        maxTotalBytes: 2048,
      }).size,
    ).toBe(2);
    const wb = new Workbook();
    wb.addWorksheet("S").addRow([42]);
    const bytes = writeWorkbookBuffer(wb);
    expect(() => readWorkbookBuffer(bytes, { maxEntries: 0 })).toThrow(ReadLimitError);
    expect(
      readWorkbookBuffer(bytes, { maxEntries: Infinity }).worksheets[0]?.getCell("A1").value,
    ).toBe(42);
    expect(DEFAULT_READ_LIMITS.maxEntries).toBe(10_000);
  });

  it.each([-1, NaN, 1.5, -Infinity])("rejects invalid configuration %s", (maxEntries) => {
    expect(() => readWorkbookBuffer(archive, { maxEntries })).toThrow(RangeError);
  });

  it("enforces an oversized declared entry before trying to inflate its corrupt payload", () => {
    const corrupt = archive.slice();
    const view = new DataView(corrupt.buffer);
    const directory = view.getUint32(corrupt.length - 22 + 16, true);
    view.setUint32(directory + 24, 5000, true);
    corrupt[35] = 255;
    expect(() => readZipArchive(corrupt, { maxEntryBytes: 1024 })).toThrow(ReadLimitError);
  });

  it("rejects a forged stored entry whose compressed size exceeds its declared original size", () => {
    // ZIP compression method 0 copies size bytes; it must not trust the smaller originalSize.
    const corrupt = archive.slice();
    const view = new DataView(corrupt.buffer);
    const directory = view.getUint32(corrupt.length - 22 + 16, true);
    view.setUint16(directory + 10, 0, true);
    view.setUint32(directory + 20, 2000, true);
    view.setUint32(directory + 24, 1, true);
    expect(() => readZipArchive(corrupt, { maxEntryBytes: 1024 })).toThrow(ReadLimitError);
  });

  it("rejects load asynchronously without replacing the receiving workbook", async () => {
    const wb = new Workbook();
    wb.addWorksheet("Original").addRow([1]);
    await expect(wb.xlsx.load(archive, { maxEntries: 0 })).rejects.toBeInstanceOf(ReadLimitError);
    expect(wb.worksheets[0]?.name).toBe("Original");
  });

  it("stops and closes an input byte stream as soon as maxInputBytes is exceeded", async () => {
    let closed = false;
    let yielded = 0;
    async function* chunks() {
      try {
        for (let i = 0; i < 10; i++) {
          yielded++;
          yield new Uint8Array(10);
        }
      } finally {
        closed = true;
      }
    }
    await expect(new Workbook().xlsx.read(chunks(), { maxInputBytes: 15 })).rejects.toMatchObject({
      code: "READ_LIMIT_EXCEEDED",
      limit: "maxInputBytes",
    });
    expect(yielded).toBe(2);
    expect(closed).toBe(true);
  });

  it("applies input limits to file reads and preserves the receiver on failure", async () => {
    const temp = await mkdtemp(join(tmpdir(), "alosha-read-limit-"));
    try {
      const path = join(temp, "oversized.xlsx");
      await writeFile(path, archive);
      const receiver = new Workbook();
      receiver.addWorksheet("Original");
      await expect(receiver.xlsx.readFile(path, { maxInputBytes: 1 })).rejects.toMatchObject({
        code: "READ_LIMIT_EXCEEDED",
        limit: "maxInputBytes",
      });
      expect(receiver.worksheets[0]?.name).toBe("Original");
    } finally {
      await rm(temp, { recursive: true, force: true });
    }
  });
});
