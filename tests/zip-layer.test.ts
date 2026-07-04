import { describe, expect, it } from "vitest";
import { readZipArchive, writeZipArchive } from "../src/zip-layer.js";

describe("zip-layer", () => {
  it("round-trips entries through write and read", () => {
    const encoder = new TextEncoder();
    const entries = new Map<string, Uint8Array>([
      ["[Content_Types].xml", encoder.encode("<Types/>")],
      ["xl/worksheets/sheet1.xml", encoder.encode("<worksheet/>")],
    ]);

    const archive = writeZipArchive(entries);
    const result = readZipArchive(archive);

    expect(result.size).toBe(entries.size);
    const decoder = new TextDecoder();
    for (const [path, data] of entries) {
      expect(result.has(path)).toBe(true);
      expect(decoder.decode(result.get(path))).toBe(decoder.decode(data));
    }
  });

  it("produces an empty map for an empty archive", () => {
    const archive = writeZipArchive(new Map());
    const result = readZipArchive(archive);
    expect(result.size).toBe(0);
  });
});
