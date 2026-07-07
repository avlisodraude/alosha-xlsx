import { describe, expect, it } from "vitest";
import type { Style } from "../../../src/model/index.js";
import { isDateFormatCode, parseStyles } from "../../../src/xlsx/read/styles.js";
import { createStyles } from "../../../src/xlsx/styles.js";

/** Add a single Style through the writer, serialize, parse it back, and return the reconstructed Style
 *  for that cellXf. The tightest possible inverse check for one style. */
function roundTrip(style: Style): Style {
  const builder = createStyles();
  const id = builder.addStyle(style);
  return parseStyles(builder.toXml()).styleFor(id);
}

/** isDateStyle for a style carrying only the given number format. */
function dateStyleOf(numFmt: string): boolean {
  const builder = createStyles();
  const id = builder.addStyle({ numFmt });
  return parseStyles(builder.toXml()).isDateStyle(id);
}

/** Add a single differential Style through the dxf writer, serialize, parse it back, and return the
 *  reconstructed Style for that dxf id. The tightest possible inverse check for one dxf. */
function dxfRoundTrip(style: Partial<Style> | undefined): Style {
  const builder = createStyles();
  const id = builder.addDxf(style);
  return parseStyles(builder.toXml()).dxfFor(id);
}

describe("parseStyles — defaults & missing part", () => {
  it("treats a missing styles part as all-default", () => {
    const table = parseStyles(undefined);
    expect(table.styleFor(0)).toEqual({});
    expect(table.styleFor(5)).toEqual({});
    expect(table.isDateStyle(5)).toBe(false);
  });

  it("returns {} for the default index 0 and out-of-range indices", () => {
    const table = parseStyles(createStyles().toXml());
    expect(table.styleFor(0)).toEqual({});
    expect(table.styleFor(999)).toEqual({});
    expect(table.styleFor(-1)).toEqual({});
    expect(table.styleFor(1.5)).toEqual({});
  });

  it("drops a General number format (it is the default)", () => {
    expect(roundTrip({ numFmt: "General" })).toEqual({});
  });
});

describe("parseStyles — number formats", () => {
  it("resolves a built-in reserved id back to its canonical code", () => {
    expect(roundTrip({ numFmt: "0.00%" })).toEqual({ numFmt: "0.00%" });
    expect(roundTrip({ numFmt: "#,##0.00" })).toEqual({ numFmt: "#,##0.00" });
  });

  it("resolves a custom (id >= 164) format code", () => {
    expect(roundTrip({ numFmt: "0.000" })).toEqual({ numFmt: "0.000" });
    expect(roundTrip({ numFmt: '"€"#,##0.00' })).toEqual({ numFmt: '"€"#,##0.00' });
  });
});

describe("parseStyles — font inversion", () => {
  it("restores a fully-specified font", () => {
    const font: Style = {
      font: {
        bold: true,
        italic: true,
        underline: "double",
        strike: true,
        size: 12,
        color: { argb: "FF112233" },
        name: "Arial",
        family: 2,
        scheme: "minor",
        vertAlign: "superscript",
      },
    };
    expect(roundTrip(font)).toEqual(font);
  });

  it("restores a plain underline as boolean true", () => {
    expect(roundTrip({ font: { underline: true } })).toEqual({ font: { underline: true } });
  });

  it("restores a theme colour", () => {
    expect(roundTrip({ font: { bold: true, color: { theme: 4 } } })).toEqual({
      font: { bold: true, color: { theme: 4 } },
    });
  });
});

describe("parseStyles — fill inversion", () => {
  it("restores a solid pattern fill with fg/bg colours", () => {
    const style: Style = {
      fill: {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFF0000" },
        bgColor: { indexed: 64 },
      },
    };
    expect(roundTrip(style)).toEqual(style);
  });

  it("restores an angle gradient fill", () => {
    const style: Style = {
      fill: {
        type: "gradient",
        gradient: "angle",
        degree: 90,
        stops: [
          { position: 0, color: { argb: "FFFF0000" } },
          { position: 1, color: { argb: "FF0000FF" } },
        ],
      },
    };
    expect(roundTrip(style)).toEqual(style);
  });

  it("restores a path gradient fill", () => {
    const style: Style = {
      fill: {
        type: "gradient",
        gradient: "path",
        center: { left: 0.5, top: 0.5 },
        stops: [{ position: 0, color: { argb: "FFFFFFFF" } }],
      },
    };
    expect(roundTrip(style)).toEqual(style);
  });
});

describe("parseStyles — border inversion", () => {
  it("restores styled sides with colours and a diagonal direction flag", () => {
    const style: Style = {
      border: {
        top: { style: "thin", color: { argb: "FF000000" } },
        left: { style: "medium" },
        diagonal: { style: "thin", up: true },
      },
    };
    expect(roundTrip(style)).toEqual(style);
  });

  it("restores a diagonal direction flag even without a diagonal style", () => {
    expect(roundTrip({ border: { top: { style: "thin" }, diagonal: { down: true } } })).toEqual({
      border: { top: { style: "thin" }, diagonal: { down: true } },
    });
  });
});

describe("parseStyles — alignment inversion", () => {
  it("restores horizontal/vertical/wrap/reading-order and vertical rotation", () => {
    const style: Style = {
      alignment: {
        horizontal: "center",
        vertical: "middle",
        textRotation: "vertical",
        wrapText: true,
        readingOrder: "rtl",
      },
    };
    expect(roundTrip(style)).toEqual(style);
  });

  it("inverts Excel's 90+ encoding back to a negative rotation", () => {
    expect(roundTrip({ alignment: { textRotation: -45 } })).toEqual({
      alignment: { textRotation: -45 },
    });
  });

  it("restores a zero rotation and an indent", () => {
    expect(roundTrip({ alignment: { textRotation: 0, indent: 2 } })).toEqual({
      alignment: { textRotation: 0, indent: 2 },
    });
  });
});

describe("parseStyles — protection inversion", () => {
  it("restores a meaningful locked=false alongside hidden=true", () => {
    expect(roundTrip({ protection: { locked: false, hidden: true } })).toEqual({
      protection: { locked: false, hidden: true },
    });
  });
});

describe("isDateStyle", () => {
  it("flags the reserved built-in date/time ids without string sniffing", () => {
    expect(dateStyleOf("mm-dd-yy")).toBe(true); // id 14
    expect(dateStyleOf("m/d/yy h:mm")).toBe(true); // id 22
    expect(dateStyleOf("mm:ss")).toBe(true); // id 45
    expect(dateStyleOf("mmss.0")).toBe(true); // id 47
  });

  it("does not flag non-date built-in formats", () => {
    expect(dateStyleOf("0.00")).toBe(false); // id 2
    expect(dateStyleOf("0.00%")).toBe(false); // id 10
    expect(dateStyleOf("#,##0.00")).toBe(false); // id 4
  });

  it("sniffs custom (id >= 164) codes for date/time tokens", () => {
    expect(dateStyleOf("yyyy-mm-dd")).toBe(true);
    expect(dateStyleOf("dd/mm/yyyy hh:mm")).toBe(true);
    expect(dateStyleOf('"Q"0')).toBe(false);
    expect(dateStyleOf("#,##0.0,,")).toBe(false);
  });

  it("is false for the default index and out-of-range indices", () => {
    const table = parseStyles(createStyles().toXml());
    expect(table.isDateStyle(0)).toBe(false);
    expect(table.isDateStyle(999)).toBe(false);
  });
});

describe("isDateFormatCode", () => {
  it("detects date/time tokens", () => {
    expect(isDateFormatCode("yyyy-mm-dd")).toBe(true);
    expect(isDateFormatCode("h:mm:ss")).toBe(true);
    expect(isDateFormatCode("[h]:mm:ss")).toBe(true); // elapsed-time bracket
    expect(isDateFormatCode("[$-409]d/m/yyyy")).toBe(true); // locale bracket then a date token
    expect(isDateFormatCode("h:mm AM/PM")).toBe(true);
  });

  it("ignores tokens inside quotes, escapes and colour/condition brackets", () => {
    expect(isDateFormatCode('0"d"')).toBe(false); // quoted literal d
    expect(isDateFormatCode("0\\d")).toBe(false); // escaped d
    expect(isDateFormatCode("[Red]#,##0")).toBe(false); // colour bracket
    expect(isDateFormatCode("[>100]0")).toBe(false); // condition bracket
  });

  it("is false for pure number and text formats", () => {
    expect(isDateFormatCode("#,##0.00")).toBe(false);
    expect(isDateFormatCode("0.00E+00")).toBe(false);
    expect(isDateFormatCode("General")).toBe(false);
    expect(isDateFormatCode("@")).toBe(false);
  });
});

describe("parseStyles — dxf pool (conditional formatting)", () => {
  it("treats a missing <dxfs> (or an empty count=0) as an empty table", () => {
    const table = parseStyles(undefined);
    expect(table.dxfFor(0)).toEqual({});

    const empty = parseStyles(createStyles().toXml());
    expect(empty.dxfFor(0)).toEqual({});
  });

  it("returns {} for an out-of-range dxf index", () => {
    const builder = createStyles();
    builder.addDxf({ font: { bold: true } });
    const table = parseStyles(builder.toXml());
    expect(table.dxfFor(-1)).toEqual({});
    expect(table.dxfFor(1.5)).toEqual({});
    expect(table.dxfFor(999)).toEqual({});
  });

  it("restores a solid-fill highlight from <patternFill><bgColor> (not fgColor)", () => {
    const style: Partial<Style> = {
      fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFF8696B" } },
    };
    expect(dxfRoundTrip(style)).toEqual(style);
  });

  it("restores a non-solid pattern's patternType alongside its colour", () => {
    const style: Partial<Style> = {
      fill: { type: "pattern", pattern: "lightGrid", bgColor: { argb: "FF00FF00" } },
    };
    expect(dxfRoundTrip(style)).toEqual(style);
  });

  it("restores a bold/italic font and a border", () => {
    const style: Partial<Style> = {
      font: { bold: true, italic: true, color: { argb: "FF000000" } },
      border: { top: { style: "thin", color: { argb: "FF000000" } } },
    };
    expect(dxfRoundTrip(style)).toEqual(style);
  });

  it("restores a numFmt by its formatCode", () => {
    expect(dxfRoundTrip({ numFmt: "0.00%" })).toEqual({ numFmt: "0.00%" });
    expect(dxfRoundTrip({ numFmt: '"€"#,##0.00' })).toEqual({ numFmt: '"€"#,##0.00' });
  });

  it("restores alignment and protection", () => {
    const style: Partial<Style> = {
      alignment: { horizontal: "center", wrapText: true },
      protection: { locked: false },
    };
    expect(dxfRoundTrip(style)).toEqual(style);
  });

  it("round-trips an empty/undefined dxf as {}", () => {
    expect(dxfRoundTrip(undefined)).toEqual({});
    expect(dxfRoundTrip({})).toEqual({});
  });

  it("dedups identical dxfs to the same id, distinct ones to different ids", () => {
    const builder = createStyles();
    const highlight: Partial<Style> = {
      fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFF8696B" } },
    };
    const id1 = builder.addDxf(highlight);
    const id2 = builder.addDxf({ ...highlight });
    const id3 = builder.addDxf({ font: { bold: true } });
    expect(id1).toBe(id2);
    expect(id3).not.toBe(id1);

    const table = parseStyles(builder.toXml());
    expect(table.dxfFor(id1)).toEqual(highlight);
    expect(table.dxfFor(id3)).toEqual({ font: { bold: true } });
  });
});

describe("parseStyles — structural symmetry (write -> read -> write)", () => {
  it("reproduces a byte-identical styles.xml and the same cellXf ids", () => {
    const styles: Style[] = [
      { font: { bold: true } },
      { font: { bold: true }, numFmt: "0.00%" }, // shares the bold font, built-in numFmt
      { numFmt: "0.000" }, // custom numFmt id 164
      { fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF0000" } } },
      {
        border: {
          top: { style: "thin", color: { argb: "FF000000" } },
          diagonal: { style: "thin", up: true },
        },
      },
      {
        alignment: { horizontal: "center", textRotation: -45, wrapText: true, readingOrder: "rtl" },
      },
      { protection: { locked: false, hidden: true } },
      {
        font: { italic: true, size: 14, color: { theme: 4 } },
        numFmt: "yyyy-mm-dd", // custom numFmt id 165
        alignment: { vertical: "middle" },
      },
    ];

    const builder = createStyles();
    const ids = styles.map((s) => builder.addStyle(s));
    const xml = builder.toXml();

    const table = parseStyles(xml);
    const rebuilt = createStyles();
    const ids2 = ids.map((id) => rebuilt.addStyle(table.styleFor(id)));

    expect(rebuilt.toXml()).toBe(xml);
    expect(ids2).toEqual(ids);
  });

  it("restores identical-font dedup so a shared font stays a single fontId", () => {
    const builder = createStyles();
    const id1 = builder.addStyle({ font: { bold: true } });
    const id2 = builder.addStyle({ font: { bold: true }, numFmt: "0.00" });
    const xml = builder.toXml();
    expect(xml).toContain('<fonts count="2">'); // default + one bold font

    const table = parseStyles(xml);
    expect(table.styleFor(id1).font).toEqual({ bold: true });
    expect(table.styleFor(id2).font).toEqual({ bold: true });
  });
});
