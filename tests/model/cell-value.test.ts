import { describe, expect, it } from "vitest";
import { InvalidCellValueError } from "../../src/errors.js";
import {
  type CellValue,
  detectType,
  effectiveType,
  fromStored,
  toCsv,
  toStored,
  toText,
} from "../../src/model/cell-value.js";
import { CellType } from "../../src/model/enums.js";

describe("detectType", () => {
  it("null -> Null", () => {
    expect(detectType(null)).toBe(CellType.Null);
  });

  it("'' -> String, not Null", () => {
    expect(detectType("")).toBe(CellType.String);
  });

  it("0 -> Number, not Null", () => {
    expect(detectType(0)).toBe(CellType.Number);
  });

  it("false -> Boolean, not Null", () => {
    expect(detectType(false)).toBe(CellType.Boolean);
  });

  it("Date -> Date", () => {
    expect(detectType(new Date())).toBe(CellType.Date);
  });

  it("{text,hyperlink} -> Hyperlink", () => {
    expect(detectType({ text: "click", hyperlink: "https://example.com" })).toBe(
      CellType.Hyperlink,
    );
  });

  it("{formula} -> Formula", () => {
    expect(detectType({ formula: "A1+A2" })).toBe(CellType.Formula);
  });

  it("{sharedFormula} -> Formula", () => {
    expect(detectType({ sharedFormula: "A1" })).toBe(CellType.Formula);
  });

  it("{richText} -> RichText", () => {
    expect(detectType({ richText: [{ text: "hi" }] })).toBe(CellType.RichText);
  });

  it("{error} -> Error", () => {
    expect(detectType({ error: "#REF!" })).toBe(CellType.Error);
  });

  it("throws InvalidCellValueError for an unsupported shape", () => {
    const bogus = { notAValidShape: true } as unknown as CellValue;
    expect(() => detectType(bogus)).toThrow(InvalidCellValueError);
  });
});

describe("toStored / fromStored round-trip", () => {
  const cases: Array<[string, CellValue]> = [
    ["null", null],
    ["number", 42],
    ["zero", 0],
    ["string", "hello"],
    ["empty string", ""],
    ["boolean true", true],
    ["boolean false", false],
    ["date", new Date("2024-01-01T00:00:00.000Z")],
    ["error", { error: "#DIV/0!" }],
    ["richText", { richText: [{ text: "hi", font: { bold: true } }] }],
    ["hyperlink", { text: "click", hyperlink: "https://example.com", tooltip: "go" }],
    ["formula", { formula: "SUM(A1:A2)", result: 3 }],
    ["shared formula", { sharedFormula: "A1", ref: "A1:A10" }],
  ];

  it.each(cases)("%s", (_label, value) => {
    const stored = toStored(value);
    expect(fromStored(stored)).toEqual(value);
  });
});

describe("fromStored edge cases", () => {
  it("Merge -> null (resolved via the master cell elsewhere)", () => {
    expect(fromStored({ type: CellType.Merge, master: "A1" })).toBeNull();
  });

  it("SharedString -> null (read-path only; resolved via the shared-strings pool)", () => {
    expect(fromStored({ type: CellType.SharedString, value: 0 })).toBeNull();
  });
});

describe("effectiveType", () => {
  it("non-formula: returns the cell's own type", () => {
    expect(effectiveType(toStored(42))).toBe(CellType.Number);
    expect(effectiveType(toStored("x"))).toBe(CellType.String);
    expect(effectiveType(toStored(null))).toBe(CellType.Null);
  });

  it("formula with a numeric result -> Number", () => {
    expect(effectiveType(toStored({ formula: "1+1", result: 2 }))).toBe(CellType.Number);
  });

  it("formula with a string result -> String", () => {
    expect(effectiveType(toStored({ formula: "A1", result: "x" }))).toBe(CellType.String);
  });

  it("formula with a boolean result -> Boolean", () => {
    expect(effectiveType(toStored({ formula: "A1=A2", result: true }))).toBe(CellType.Boolean);
  });

  it("formula with a date result -> Date", () => {
    expect(effectiveType(toStored({ formula: "TODAY()", result: new Date() }))).toBe(CellType.Date);
  });

  it("formula with an error result -> Error", () => {
    expect(effectiveType(toStored({ formula: "1/0", result: { error: "#DIV/0!" } }))).toBe(
      CellType.Error,
    );
  });

  it("formula with no cached result -> Null", () => {
    expect(effectiveType(toStored({ formula: "A1" }))).toBe(CellType.Null);
  });
});

describe("toText", () => {
  it("Null / Merge / SharedString -> ''", () => {
    expect(toText({ type: CellType.Null })).toBe("");
    expect(toText({ type: CellType.Merge, master: "A1" })).toBe("");
    expect(toText({ type: CellType.SharedString, value: 0 })).toBe("");
  });

  it("renders Number/String/Boolean/Date/Error", () => {
    expect(toText(toStored(42))).toBe("42");
    expect(toText(toStored("hi"))).toBe("hi");
    expect(toText(toStored(true))).toBe("TRUE");
    expect(toText(toStored(false))).toBe("FALSE");
    expect(toText(toStored({ error: "#N/A" }))).toBe("#N/A");
  });

  it("RichText concatenates run text", () => {
    expect(toText(toStored({ richText: [{ text: "Hello, " }, { text: "world" }] }))).toBe(
      "Hello, world",
    );
  });

  it("Hyperlink uses the display text, not the target", () => {
    expect(toText(toStored({ text: "click here", hyperlink: "https://x" }))).toBe("click here");
  });

  it("Formula renders its cached result, or '' if none", () => {
    expect(toText(toStored({ formula: "1+1", result: 2 }))).toBe("2");
    expect(toText(toStored({ formula: "A1" }))).toBe("");
  });
});

describe("toCsv", () => {
  it("passes plain text through unquoted", () => {
    expect(toCsv(toStored("plain"))).toBe("plain");
    expect(toCsv(toStored(42))).toBe("42");
  });

  it("quotes and escapes strings containing commas, quotes, or newlines", () => {
    expect(toCsv(toStored("a,b"))).toBe('"a,b"');
    expect(toCsv(toStored('say "hi"'))).toBe('"say ""hi"""');
    expect(toCsv(toStored("line1\nline2"))).toBe('"line1\nline2"');
  });

  it("renders a formula's result, CSV-escaped", () => {
    expect(toCsv(toStored({ formula: "A1", result: "a,b" }))).toBe('"a,b"');
  });

  it("'' for null/merge", () => {
    expect(toCsv({ type: CellType.Null })).toBe("");
    expect(toCsv({ type: CellType.Merge, master: "A1" })).toBe("");
  });
});
