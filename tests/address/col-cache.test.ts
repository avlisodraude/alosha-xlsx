import { describe, expect, it } from "vitest";
import {
  colLetterToNumber,
  colNumberToLetter,
  DECODE_ADDRESS_CACHE_MAX,
  decodeAddress,
  decodeEx,
  encodeAddress,
  validateAddress,
} from "../../src/address/col-cache.js";
import { InvalidAddressError } from "../../src/errors.js";

describe("colLetterToNumber / colNumberToLetter", () => {
  it.each([
    ["A", 1],
    ["Z", 26],
    ["AA", 27],
    ["XFD", 16384],
  ])("%s <-> %i", (letter, num) => {
    expect(colLetterToNumber(letter)).toBe(num);
    expect(colNumberToLetter(num)).toBe(letter);
  });

  it("is case-insensitive on input", () => {
    expect(colLetterToNumber("a")).toBe(1);
    expect(colLetterToNumber("aa")).toBe(27);
  });

  it.each(["", "1", "A1", "@", "A-"])("colLetterToNumber(%j) throws InvalidAddressError", (v) => {
    expect(() => colLetterToNumber(v)).toThrow(InvalidAddressError);
  });

  it.each([0, -1, 1.5, Number.NaN])("colNumberToLetter(%j) throws InvalidAddressError", (v) => {
    expect(() => colNumberToLetter(v)).toThrow(InvalidAddressError);
  });

  it("returns stable results across repeated calls (memoized)", () => {
    expect(colLetterToNumber("BQ")).toBe(colLetterToNumber("BQ"));
    expect(colNumberToLetter(69)).toBe(colNumberToLetter(69));
  });
});

describe("encodeAddress", () => {
  it("(row=1,col=2) -> 'B1'", () => {
    expect(encodeAddress(1, 2)).toBe("B1");
  });

  it("supports multi-letter columns and large rows", () => {
    expect(encodeAddress(1_048_576, 16384)).toBe("XFD1048576");
  });
});

describe("decodeAddress", () => {
  it("'B1' -> { row:1, col:2, address:'B1' }", () => {
    expect(decodeAddress("B1")).toEqual({ row: 1, col: 2, address: "B1" });
  });

  it("normalizes casing of the returned address", () => {
    expect(decodeAddress("b1")).toEqual({ row: 1, col: 2, address: "B1" });
  });

  it("supports multi-letter columns and rows up to 1,048,576", () => {
    expect(decodeAddress("XFD1048576")).toEqual({
      row: 1_048_576,
      col: 16384,
      address: "XFD1048576",
    });
  });

  it.each([
    "",
    "A0",
    "1A",
    "@",
    "$A$1",
    "A",
    "1",
    "XFE1",
    "A1048577",
  ])("decodeAddress(%j) throws InvalidAddressError", (v) => {
    expect(() => decodeAddress(v)).toThrow(InvalidAddressError);
  });

  it("returns a stable (cached) result for repeated calls", () => {
    const first = decodeAddress("Q17");
    const second = decodeAddress("Q17");
    expect(second).toBe(first);
    expect(second).toEqual({ row: 17, col: 17, address: "Q17" });
  });

  it("encode/decode round-trip fuzz", () => {
    for (let i = 0; i < 200; i++) {
      const row = 1 + Math.floor(Math.random() * 1_048_576);
      const col = 1 + Math.floor(Math.random() * 16384);
      const address = encodeAddress(row, col);
      expect(decodeAddress(address)).toEqual({ row, col, address });
    }
  });

  it("caps its memo (FIFO eviction) so bulk decodes don't leak", () => {
    // Seed one address, then decode more distinct addresses than the cap so the seed is evicted.
    const seeded = decodeAddress("A1");
    for (let row = 2; row <= DECODE_ADDRESS_CACHE_MAX + 2; row++) {
      decodeAddress(`A${row}`);
    }
    // Still correct after eviction, but no longer the same cached instance — proof it was dropped
    // rather than retained forever (an unbounded cache would return the original object).
    const afterEviction = decodeAddress("A1");
    expect(afterEviction).toEqual({ row: 1, col: 1, address: "A1" });
    expect(afterEviction).not.toBe(seeded);
  });
});

describe("decodeEx", () => {
  it("handles absolute '$A$1'", () => {
    expect(decodeEx("$A$1")).toEqual({
      row: 1,
      col: 1,
      address: "A1",
      $col$row: "$A$1",
    });
  });

  it("handles a quoted sheet name with a space", () => {
    expect(decodeEx("'Sheet 1'!B2")).toEqual({
      row: 2,
      col: 2,
      address: "B2",
      $col$row: "$B$2",
      sheetName: "Sheet 1",
    });
  });

  it("handles an unquoted sheet name", () => {
    expect(decodeEx("Sheet1!C3")).toEqual({
      row: 3,
      col: 3,
      address: "C3",
      $col$row: "$C$3",
      sheetName: "Sheet1",
    });
  });

  it("handles a plain relative address (no sheet, no $)", () => {
    const result = decodeEx("D4");
    expect(result.sheetName).toBeUndefined();
    expect(result.$col$row).toBe("$D$4");
  });

  it("handles mixed absolute/relative markers", () => {
    expect(decodeEx("$A1").address).toBe("A1");
    expect(decodeEx("A$1").address).toBe("A1");
  });

  it("throws InvalidAddressError for malformed input", () => {
    expect(() => decodeEx("not-an-address")).toThrow(InvalidAddressError);
  });
});

describe("validateAddress", () => {
  it("does not throw for well-formed addresses", () => {
    expect(() => validateAddress("A1")).not.toThrow();
    expect(() => validateAddress("XFD1048576")).not.toThrow();
  });

  it.each(["", "A0", "1A", "@", "$A$1"])("throws InvalidAddressError for %j", (v) => {
    expect(() => validateAddress(v)).toThrow(InvalidAddressError);
  });
});
