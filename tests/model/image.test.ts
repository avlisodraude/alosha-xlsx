import { describe, expect, it } from "vitest";
import { InvalidImageAnchorError, InvalidRangeError } from "../../src/errors.js";
import { type ImageAnchor, ImageRegistry, resolveImageAnchor } from "../../src/model/image.js";
import { Workbook } from "../../src/model/workbook.js";
import type { Worksheet } from "../../src/model/worksheet.js";

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
const JPEG = new Uint8Array([0xff, 0xd8, 0xff, 0xe0, 0x00]);

describe("ImageRegistry", () => {
  it("assigns 0-based ids in insertion order", () => {
    const reg = new ImageRegistry();
    expect(reg.add({ buffer: PNG, extension: "png" })).toBe(0);
    expect(reg.add({ buffer: JPEG, extension: "jpeg" })).toBe(1);
    expect(reg.size).toBe(2);
  });

  it("does not dedup identical buffers (v1)", () => {
    const reg = new ImageRegistry();
    const a = reg.add({ buffer: PNG, extension: "png" });
    const b = reg.add({ buffer: PNG, extension: "png" });
    expect(a).toBe(0);
    expect(b).toBe(1);
  });

  it("returns the interned bytes and format by id, undefined otherwise", () => {
    const reg = new ImageRegistry();
    const id = reg.add({ buffer: PNG, extension: "png" });
    expect(reg.get(id)).toEqual({ buffer: PNG, extension: "png" });
    expect(reg.get(99)).toBeUndefined();
  });
});

describe("Workbook image registry wiring", () => {
  it("addImage interns bytes and returns a numeric id; getImage fetches them back", () => {
    const wb = new Workbook();
    const id = wb.addImage({ buffer: PNG, extension: "png" });
    expect(id).toBe(0);
    expect(wb.getImage(id)).toEqual({ buffer: PNG, extension: "png" });
  });
});

describe("resolveImageAnchor", () => {
  it("decodes a two-cell range string into 0-based tl/br", () => {
    expect(resolveImageAnchor("B2:D6")).toEqual({
      tl: { col: 1, row: 1 },
      br: { col: 3, row: 5 },
    });
  });

  it("decodes a single-cell string as a degenerate two-cell box", () => {
    expect(resolveImageAnchor("A1")).toEqual({
      tl: { col: 0, row: 0 },
      br: { col: 0, row: 0 },
    });
  });

  it("normalizes an unordered range string (Range sorts corners)", () => {
    expect(resolveImageAnchor("D6:B2")).toEqual({
      tl: { col: 1, row: 1 },
      br: { col: 3, row: 5 },
    });
  });

  it("passes through an explicit two-cell anchor and preserves offsets", () => {
    const anchor: ImageAnchor = {
      tl: { col: 1, row: 1, nativeColOff: 100 },
      br: { col: 3, row: 5 },
    };
    const resolved = resolveImageAnchor(anchor);
    expect(resolved).toEqual({
      tl: { col: 1, row: 1, nativeColOff: 100 },
      br: { col: 3, row: 5 },
    });
    // Deep-copied, not aliased.
    expect(resolved.tl).not.toBe(anchor.tl);
  });

  it("passes through a one-cell + ext anchor", () => {
    expect(
      resolveImageAnchor({ tl: { col: 2, row: 3 }, ext: { width: 200, height: 120 } }),
    ).toEqual({ tl: { col: 2, row: 3 }, ext: { width: 200, height: 120 } });
  });

  it("throws InvalidRangeError for a malformed range string", () => {
    expect(() => resolveImageAnchor("A1:")).toThrow(InvalidRangeError);
    expect(() => resolveImageAnchor("@@@")).toThrow(InvalidRangeError);
  });

  it("throws InvalidImageAnchorError for a malformed object anchor", () => {
    // Neither br nor ext.
    expect(() => resolveImageAnchor({ tl: { col: 0, row: 0 } } as unknown as ImageAnchor)).toThrow(
      InvalidImageAnchorError,
    );
    // Non-integer / negative position.
    expect(() => resolveImageAnchor({ tl: { col: -1, row: 0 }, br: { col: 1, row: 1 } })).toThrow(
      InvalidImageAnchorError,
    );
    // Non-positive extent.
    expect(() =>
      resolveImageAnchor({ tl: { col: 0, row: 0 }, ext: { width: 0, height: 10 } }),
    ).toThrow(InvalidImageAnchorError);
    // Not an object at all.
    expect(() => resolveImageAnchor(42 as unknown as ImageAnchor)).toThrow(InvalidImageAnchorError);
  });
});

describe("Worksheet.addImage", () => {
  let ws: Worksheet;

  it("starts with no images", () => {
    ws = new Workbook().addWorksheet("Sheet1");
    expect(ws.images).toEqual([]);
  });

  it("stores placed images with resolved anchors, in insertion order", () => {
    ws = new Workbook().addWorksheet("Sheet1");
    const first = ws.addImage(0, "B2:D6");
    ws.addImage(1, { tl: { col: 0, row: 0 }, ext: { width: 64, height: 64 } });
    expect(first).toEqual({
      imageId: 0,
      range: { tl: { col: 1, row: 1 }, br: { col: 3, row: 5 } },
    });
    expect(ws.images).toEqual([
      { imageId: 0, range: { tl: { col: 1, row: 1 }, br: { col: 3, row: 5 } } },
      { imageId: 1, range: { tl: { col: 0, row: 0 }, ext: { width: 64, height: 64 } } },
    ]);
  });

  it("propagates the typed error from a bad anchor", () => {
    ws = new Workbook().addWorksheet("Sheet1");
    expect(() => ws.addImage(0, "nope!")).toThrow(InvalidRangeError);
    expect(() => ws.addImage(0, {} as unknown as ImageAnchor)).toThrow(InvalidImageAnchorError);
  });

  it("images getter returns a defensive copy", () => {
    ws = new Workbook().addWorksheet("Sheet1");
    ws.addImage(0, "A1:B2");
    const snapshot = ws.images;
    (snapshot[0] as { imageId: number }).imageId = 999;
    expect(ws.images[0]?.imageId).toBe(0);
  });
});

describe("model round-trip", () => {
  it("round-trips the workbook registry and per-sheet anchors", () => {
    const wb = new Workbook();
    const png = wb.addImage({ buffer: PNG, extension: "png" });
    const jpeg = wb.addImage({ buffer: JPEG, extension: "jpeg" });
    const ws = wb.addWorksheet("Sheet1");
    ws.addImage(png, "B2:D6");
    ws.addImage(jpeg, { tl: { col: 0, row: 0 }, ext: { width: 100, height: 50 } });

    const clone = new Workbook();
    clone.model = wb.model;

    expect(clone.getImage(png)).toEqual({ buffer: PNG, extension: "png" });
    expect(clone.getImage(jpeg)).toEqual({ buffer: JPEG, extension: "jpeg" });
    const clonedWs = clone.getWorksheet("Sheet1");
    expect(clonedWs?.images).toEqual(ws.images);
  });

  it("omits image keys from the model when there are none", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    expect(wb.model.images).toBeUndefined();
    expect(ws.model.images).toBeUndefined();
  });

  it("re-serializing a round-tripped model is stable", () => {
    const wb = new Workbook();
    wb.addImage({ buffer: PNG, extension: "png" });
    const ws = wb.addWorksheet("Sheet1");
    ws.addImage(0, "B2:D6");

    const clone = new Workbook();
    clone.model = wb.model;
    expect(clone.model).toEqual(wb.model);
  });
});
