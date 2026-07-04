import { describe, expect, it } from "vitest";
import { Range } from "../../src/address/range.js";
import { InvalidRangeError } from "../../src/errors.js";

describe("Range construction", () => {
  it("defaults to a single-cell box at A1 with no source", () => {
    const r = new Range();
    expect([r.top, r.left, r.bottom, r.right]).toEqual([1, 1, 1, 1]);
  });

  it("from a single-cell string 'B2'", () => {
    const r = new Range("B2");
    expect([r.top, r.left, r.bottom, r.right]).toEqual([2, 2, 2, 2]);
  });

  it("from a range string 'A1:B2'", () => {
    const r = new Range("A1:B2");
    expect([r.top, r.left, r.bottom, r.right]).toEqual([1, 1, 2, 2]);
  });

  it("from a ['A1','B2'] tuple", () => {
    const r = new Range(["A1", "B2"]);
    expect([r.top, r.left, r.bottom, r.right]).toEqual([1, 1, 2, 2]);
  });

  it("from a [top,left,bottom,right] tuple", () => {
    const r = new Range([1, 1, 2, 2]);
    expect([r.top, r.left, r.bottom, r.right]).toEqual([1, 1, 2, 2]);
  });

  it("from the (top,left,bottom,right) constructor overload", () => {
    const r = new Range(1, 1, 2, 2);
    expect([r.top, r.left, r.bottom, r.right]).toEqual([1, 1, 2, 2]);
  });

  it("from another Range (copy)", () => {
    const src = new Range("A1:C3");
    const r = new Range(src);
    expect([r.top, r.left, r.bottom, r.right]).toEqual([1, 1, 3, 3]);
    expect(r).not.toBe(src);
  });

  it("from a Decoded object (single cell)", () => {
    const r = new Range({ row: 5, col: 6, address: "F5" });
    expect([r.top, r.left, r.bottom, r.right]).toEqual([5, 6, 5, 6]);
  });

  it("captures sheetName from a sheet-qualified string", () => {
    const r = new Range("'Sheet 1'!A1:B2");
    expect(r.sheetName).toBe("Sheet 1");
  });

  it.each([
    ["B2:A1", [1, 1, 2, 2]],
    [[2, 2, 1, 1] as [number, number, number, number], [1, 1, 2, 2]],
    [["B2", "A1"] as [string, string], [1, 1, 2, 2]],
  ])("normalizes unordered input %j", (source, expected) => {
    const r = new Range(source);
    expect([r.top, r.left, r.bottom, r.right]).toEqual(expected);
  });

  it.each([
    "A1:",
    "not-a-range",
    "",
    "A0:B1",
  ])("throws InvalidRangeError for malformed string %j", (v) => {
    expect(() => new Range(v)).toThrow(InvalidRangeError);
  });

  it("throws InvalidRangeError for a malformed tuple", () => {
    expect(() => new Range(["A1", "not-an-address"])).toThrow(InvalidRangeError);
  });

  it("throws InvalidRangeError for an unsupported source shape", () => {
    // biome-ignore lint/suspicious/noExplicitAny: intentionally invalid input for the throw path
    expect(() => new Range(42 as any)).toThrow(InvalidRangeError);
  });
});

describe("Range getters", () => {
  it("tl / br / range / count", () => {
    const r = new Range("A1:C2");
    expect(r.tl).toBe("A1");
    expect(r.br).toBe("C2");
    expect(r.range).toBe("A1:C2");
    expect(r.count).toBe(6);
  });

  it("count for a single cell is 1", () => {
    expect(new Range("D4").count).toBe(1);
  });
});

describe("Range#expand", () => {
  it("grows the box to include the given bounds", () => {
    const r = new Range("B2:C3");
    r.expand(1, 1, 4, 4);
    expect([r.top, r.left, r.bottom, r.right]).toEqual([1, 1, 4, 4]);
  });

  it("does not shrink when given a smaller box", () => {
    const r = new Range("B2:D4");
    r.expand(3, 3, 3, 3);
    expect([r.top, r.left, r.bottom, r.right]).toEqual([2, 2, 4, 4]);
  });
});

describe("Range#contains / Range#intersects truth table", () => {
  const r = new Range("B2:D4"); // rows 2-4, cols 2-4

  it.each([
    ["B2", true],
    ["D4", true],
    ["C3", true],
    ["A1", false],
    ["E5", false],
    ["B1", false],
    ["A2", false],
  ])("contains(%s) === %s", (address, expected) => {
    expect(r.contains(address)).toBe(expected);
  });

  it.each([
    ["C3:C3", true], // fully inside
    ["A1:B2", true], // overlaps at corner B2
    ["D4:F6", true], // overlaps at corner D4
    ["A1:A10", false], // disjoint column
    ["E1:F10", false], // disjoint column, to the right
    ["A5:F5", false], // disjoint row, below
    ["B2:D4", true], // identical box
  ])("intersects(%s) === %s", (other, expected) => {
    expect(r.intersects(new Range(other))).toBe(expected);
  });
});

describe("Range#forEachAddress", () => {
  it("visits `count` cells in row-major order", () => {
    const r = new Range("A1:B2");
    const visited: Array<{ address: string; row: number; col: number }> = [];
    r.forEachAddress((address, row, col) => visited.push({ address, row, col }));

    expect(visited).toHaveLength(r.count);
    expect(visited).toEqual([
      { address: "A1", row: 1, col: 1 },
      { address: "B1", row: 1, col: 2 },
      { address: "A2", row: 2, col: 1 },
      { address: "B2", row: 2, col: 2 },
    ]);
  });

  it("visits exactly one cell for a single-cell range", () => {
    const visited: string[] = [];
    new Range("F5").forEachAddress((address) => visited.push(address));
    expect(visited).toEqual(["F5"]);
  });
});
