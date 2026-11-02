import { describe, expect, it } from "vitest";
import { decodeAddress, decodeEx, encodeAddress, Range, Workbook } from "../../src/index.js";

describe("address cache and quoted sheet references", () => {
  it("prevents runtime mutation of cached coordinates", () => {
    const decoded = decodeAddress("B10");
    expect(Object.isFrozen(decoded)).toBe(true);
    for (const [key, value] of [
      ["row", 999],
      ["col", 999],
      ["address", "Z999"],
    ] as const) {
      expect(Reflect.set(decoded, key, value)).toBe(false);
    }
    expect(decodeAddress("B10")).toEqual({ row: 10, col: 2, address: "B10" });
    expect(new Workbook().addWorksheet("S").getCell("B10").row).toBe(10);
  });

  it.each([
    "O'Brien",
    "Two ' quotes ' here",
    "Ends'",
    "'Starts",
    "Bang! Name",
    "A''B",
  ])("decodes sheet name %j", (name) => {
    const escaped = name.replace(/'/g, "''");
    expect(decodeEx(`'${escaped}'!$XFD$1048576`)).toMatchObject({
      sheetName: name,
      row: 1048576,
      col: 16384,
    });
    expect(new Range(`'${escaped}'!A1:B2`).sheetName).toBe(name);
  });

  it("round-trips a deterministic grid sample through all address forms", () => {
    for (let i = 1; i <= 512; i++) {
      const row = 1 + ((i * 7919) % 1048576);
      const col = 1 + ((i * 97) % 16384);
      const address = encodeAddress(row, col);
      const decoded = decodeAddress(address);
      expect(decodeEx(`'O''Brien'!${address}`)).toMatchObject({ row, col, sheetName: "O'Brien" });
      expect(new Range(decoded).tl).toBe(address);
      expect(Reflect.set(decoded, "row", 1)).toBe(false);
      expect(decodeAddress(address).row).toBe(row);
    }
  });
});
