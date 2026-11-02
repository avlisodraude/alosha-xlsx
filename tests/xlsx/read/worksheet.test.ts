import { describe, expect, it } from "vitest";
import {
  InvalidConditionalFormattingError,
  InvalidPaneError,
  isAloshaXlsxError,
} from "../../../src/errors.js";
import {
  CellType,
  type RichText,
  type Style,
  Workbook,
  type Worksheet,
} from "../../../src/model/index.js";
import { serialToDate } from "../../../src/xlsx/date.js";
import { REL_TYPE } from "../../../src/xlsx/parts/constants.js";
import { renderWorksheet } from "../../../src/xlsx/parts/worksheet.js";
import { parseSharedStrings } from "../../../src/xlsx/read/shared-strings.js";
import { parseStyles, type StyleTable } from "../../../src/xlsx/read/styles.js";
import { parseWorksheet, type WorksheetReadContext } from "../../../src/xlsx/read/worksheet.js";
import { createSharedStrings } from "../../../src/xlsx/shared-strings.js";
import { createStyles } from "../../../src/xlsx/styles.js";

// --- test scaffolding --------------------------------------------------------------------------------

function sheet(): Worksheet {
  return new Workbook().addWorksheet("S");
}

/** A StyleTable stub: `styleFor` returns whatever `styles[index]` holds (default {}); `isDateStyle`
 *  is true for the listed indices; `dxfFor` returns whatever `dxfs[index]` holds (default {}). Lets
 *  grid tests exercise date-vs-number without a real styles.xml. */
function stubStyles(opts?: {
  dateIndices?: number[];
  styles?: Record<number, Style>;
  dxfs?: Record<number, Style>;
}): StyleTable {
  const dates = new Set(opts?.dateIndices ?? []);
  const map = opts?.styles ?? {};
  const dxfMap = opts?.dxfs ?? {};
  return {
    styleFor: (i) => map[i] ?? {},
    isDateStyle: (i) => dates.has(i),
    dxfFor: (i) => dxfMap[i] ?? {},
  };
}

function ctx(overrides?: Partial<WorksheetReadContext>): WorksheetReadContext {
  return {
    sharedStrings: [],
    styles: stubStyles(),
    date1904: false,
    ...overrides,
  };
}

/** Wrap grid markup in a minimal worksheet element. */
function ws(inner: string): string {
  return `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${inner}</worksheet>`;
}

// --- value decoding by `t` ---------------------------------------------------------------------------

describe("parseWorksheet — value decoding", () => {
  it("decodes a bare number (t absent)", () => {
    const s = sheet();
    parseWorksheet(s, ws(`<sheetData><row r="1"><c r="A1"><v>42</v></c></row></sheetData>`), ctx());
    expect(s.getCell("A1").value).toBe(42);
    expect(s.getCell("A1").type).toBe(CellType.Number);
  });

  it("decodes a shared string (t=s) via the pool index", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(`<sheetData><row r="1"><c r="A1" t="s"><v>1</v></c></row></sheetData>`),
      ctx({ sharedStrings: ["zero", "one"] }),
    );
    expect(s.getCell("A1").value).toBe("one");
  });

  it("decodes a shared rich-text entry", () => {
    const rich: RichText = { richText: [{ text: "a" }, { text: "b", font: { bold: true } }] };
    const s = sheet();
    parseWorksheet(
      s,
      ws(`<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>`),
      ctx({ sharedStrings: [rich] }),
    );
    expect(s.getCell("A1").value).toEqual(rich);
  });

  it("does not alias one RichText object across cells sharing a shared-string index", () => {
    const rich: RichText = { richText: [{ text: "a", font: { bold: true } }, { text: "b" }] };
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>0</v></c></row></sheetData>`,
      ),
      ctx({ sharedStrings: [rich] }),
    );
    const a = s.getCell("A1").value as RichText;
    const b = s.getCell("B1").value as RichText;
    expect(a).toEqual(rich);
    expect(b).toEqual(rich);
    // Distinct objects (and distinct runs), so editing one cell can't mutate the other or the pool.
    expect(a).not.toBe(b);
    expect(a.richText[0]).not.toBe(b.richText[0]);
    (a.richText[0] as { text: string }).text = "MUTATED";
    expect(b.richText[0]?.text).toBe("a");
    expect(rich.richText[0]?.text).toBe("a");
  });

  it("decodes an inline string (t=inlineStr)", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Inline</t></is></c></row></sheetData>`,
      ),
      ctx(),
    );
    expect(s.getCell("A1").value).toBe("Inline");
  });

  it("decodes inline rich-text runs", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<sheetData><row r="1"><c r="A1" t="inlineStr"><is><r><rPr><b/></rPr><t>x</t></r><r><t>y</t></r></is></c></row></sheetData>`,
      ),
      ctx(),
    );
    expect(s.getCell("A1").value).toEqual({
      richText: [{ text: "x", font: { bold: true } }, { text: "y" }],
    });
  });

  it("decodes a boolean (t=b) both ways", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<sheetData><row r="1"><c r="A1" t="b"><v>1</v></c><c r="B1" t="b"><v>0</v></c></row></sheetData>`,
      ),
      ctx(),
    );
    expect(s.getCell("A1").value).toBe(true);
    expect(s.getCell("B1").value).toBe(false);
  });

  it("decodes an error (t=e)", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(`<sheetData><row r="1"><c r="A1" t="e"><v>#REF!</v></c></row></sheetData>`),
      ctx(),
    );
    expect(s.getCell("A1").value).toEqual({ error: "#REF!" });
  });

  it("decodes a bare string result (t=str)", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(`<sheetData><row r="1"><c r="A1" t="str"><v>hi</v></c></row></sheetData>`),
      ctx(),
    );
    expect(s.getCell("A1").value).toBe("hi");
  });

  it("promotes a number to a Date when the style is a date format", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(`<sheetData><row r="1"><c r="A1" s="5"><v>44927</v></c></row></sheetData>`),
      ctx({ styles: stubStyles({ dateIndices: [5] }) }),
    );
    const value = s.getCell("A1").value;
    expect(value).toBeInstanceOf(Date);
    expect((value as Date).getTime()).toBe(serialToDate(44927, false).getTime());
  });

  it("keeps a serial as a number when the style is not a date format", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(`<sheetData><row r="1"><c r="A1" s="5"><v>44927</v></c></row></sheetData>`),
      ctx({ styles: stubStyles({ dateIndices: [] }) }),
    );
    expect(s.getCell("A1").value).toBe(44927);
  });

  it("honours date1904 when resolving a date serial", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(`<sheetData><row r="1"><c r="A1" s="1"><v>1000</v></c></row></sheetData>`),
      ctx({ styles: stubStyles({ dateIndices: [1] }), date1904: true }),
    );
    expect((s.getCell("A1").value as Date).getTime()).toBe(serialToDate(1000, true).getTime());
  });
});

// --- formulas ----------------------------------------------------------------------------------------

describe("parseWorksheet — formulas", () => {
  it("reads a formula with a cached number result", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(`<sheetData><row r="1"><c r="A1"><f>A2+1</f><v>43</v></c></row></sheetData>`),
      ctx(),
    );
    const cell = s.getCell("A1");
    expect(cell.type).toBe(CellType.Formula);
    expect(cell.formula).toBe("A2+1");
    expect(cell.result).toBe(43);
  });

  it("reads a formula with a cached string result (t=str)", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(`<sheetData><row r="1"><c r="A1" t="str"><f>UPPER(B1)</f><v>HI</v></c></row></sheetData>`),
      ctx(),
    );
    expect(s.getCell("A1").formula).toBe("UPPER(B1)");
    expect(s.getCell("A1").result).toBe("HI");
  });

  it("reads a formula with a cached date result via the date style", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(`<sheetData><row r="1"><c r="A1" s="3"><f>TODAY()</f><v>44927</v></c></row></sheetData>`),
      ctx({ styles: stubStyles({ dateIndices: [3] }) }),
    );
    expect(s.getCell("A1").result).toBeInstanceOf(Date);
  });

  it("reads a formula with no cached value", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(`<sheetData><row r="1"><c r="A1"><f>A2+1</f></c></row></sheetData>`),
      ctx(),
    );
    expect(s.getCell("A1").formula).toBe("A2+1");
    expect(s.getCell("A1").result).toBeUndefined();
  });
});

// --- errors ------------------------------------------------------------------------------------------

describe("parseWorksheet — malformed input", () => {
  it("throws a typed error for a non-numeric boolean value", () => {
    const s = sheet();
    let caught: unknown;
    try {
      parseWorksheet(
        s,
        ws(`<sheetData><row r="1"><c r="A1" t="b"><v>maybe</v></c></row></sheetData>`),
        ctx(),
      );
    } catch (e) {
      caught = e;
    }
    expect(isAloshaXlsxError(caught)).toBe(true);
  });
});

// --- cols, rows, styles, placement -------------------------------------------------------------------

describe("parseWorksheet — columns", () => {
  it("expands <col> ranges into per-column width/hidden/outline", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<cols><col min="1" max="2" width="20" customWidth="1"/><col min="3" max="3" hidden="1" outlineLevel="2"/></cols><sheetData/>`,
      ),
      ctx(),
    );
    expect(s.getColumn(1).width).toBe(20);
    expect(s.getColumn(2).width).toBe(20);
    expect(s.getColumn(3).hidden).toBe(true);
    expect(s.getColumn(3).outlineLevel).toBe(2);
  });

  it("applies a column style from its style index", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(`<cols><col min="1" max="1" style="7"/></cols><sheetData/>`),
      ctx({ styles: stubStyles({ styles: { 7: { font: { bold: true } } } }) }),
    );
    expect(s.getColumn(1).style).toEqual({ font: { bold: true } });
  });
});

describe("parseWorksheet — rows", () => {
  it("reads row height/hidden/outline attributes", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<sheetData><row r="2" ht="30" customHeight="1" hidden="1" outlineLevel="1"><c r="A2"><v>1</v></c></row></sheetData>`,
      ),
      ctx(),
    );
    const row = s.getRow(2);
    expect(row.height).toBe(30);
    expect(row.hidden).toBe(true);
    expect(row.outlineLevel).toBe(1);
  });

  it("applies a row style from its s attribute", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(`<sheetData><row r="1" s="9" customFormat="1"><c r="A1"><v>1</v></c></row></sheetData>`),
      ctx({ styles: stubStyles({ styles: { 9: { font: { italic: true } } } }) }),
    );
    expect(s.getRow(1).style).toEqual({ font: { italic: true } });
  });
});

describe("parseWorksheet — cell placement & styling", () => {
  it("skips empty, unstyled cells but keeps their neighbours", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(`<sheetData><row r="1"><c r="A1"/><c r="B1"><v>5</v></c></row></sheetData>`),
      ctx(),
    );
    expect(s.findCell("A1")).toBeUndefined();
    expect(s.getCell("B1").value).toBe(5);
  });

  it("keeps a styled-but-empty cell and hangs its style on it", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(`<sheetData><row r="1"><c r="A1" s="2"/></row></sheetData>`),
      ctx({ styles: stubStyles({ styles: { 2: { font: { bold: true } } } }) }),
    );
    const cell = s.findCell("A1");
    expect(cell).toBeDefined();
    expect(cell?.type).toBe(CellType.Null);
    expect(cell?.style).toEqual({ font: { bold: true } });
  });

  it("places cells without an r attribute in successive columns", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(`<sheetData><row r="1"><c><v>1</v></c><c><v>2</v></c></row></sheetData>`),
      ctx(),
    );
    expect(s.getCell(1, 1).value).toBe(1);
    expect(s.getCell(1, 2).value).toBe(2);
  });

  it("hangs a non-default cell style straight on the cell", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(`<sheetData><row r="1"><c r="A1" s="4"><v>7</v></c></row></sheetData>`),
      ctx({ styles: stubStyles({ styles: { 4: { numFmt: "0.00" } } }) }),
    );
    expect(s.getCell("A1").numFmt).toBe("0.00");
  });
});

// --- merges ------------------------------------------------------------------------------------------

describe("parseWorksheet — merges", () => {
  it("applies merge ranges after cells so the master keeps its value", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Merged</t></is></c></row></sheetData><mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>`,
      ),
      ctx(),
    );
    expect(s.getCell("A1").value).toBe("Merged");
    expect(s.getCell("B1").isMerged).toBe(true);
    expect(s.getCell("B1").master.address).toBe("A1");
  });
});

// --- full writer -> reader round-trip ----------------------------------------------------------------

describe("parseWorksheet — round-trip through the writer", () => {
  function fixture(): Worksheet {
    const src = new Workbook().addWorksheet("Data");
    src.getCell("A1").value = "Hello";
    src.getCell("B1").value = "Hello"; // dedups to the same shared-string index
    src.getCell("A2").value = 42;
    src.getCell("A2").numFmt = "0.00";
    src.getCell("B2").value = true;
    src.getCell("A3").value = new Date("2026-01-15T00:00:00.000Z");
    src.getCell("B3").value = { error: "#REF!" };
    src.getCell("A4").value = { formula: "A2+1", result: 43 };
    src.getColumn(1).width = 20;
    src.getCell("A5").value = "Merged";
    src.mergeCells("A5:B5");
    return src;
  }

  function readBack(useSharedStrings: boolean): Worksheet {
    const styles = createStyles();
    const sharedStrings = createSharedStrings();
    const { xml } = renderWorksheet(fixture(), {
      styles,
      sharedStrings,
      useSharedStrings,
      date1904: false,
    });

    const table = parseStyles(styles.toXml());
    const sst = parseSharedStrings(sharedStrings.toXml() ?? "");

    const dst = new Workbook().addWorksheet("Data");
    parseWorksheet(dst, xml, { sharedStrings: sst, styles: table, date1904: false });
    return dst;
  }

  it("reproduces every value, style, merge and column width (shared strings)", () => {
    const dst = readBack(true);
    expect(dst.getCell("A1").value).toBe("Hello");
    expect(dst.getCell("B1").value).toBe("Hello");
    expect(dst.getCell("A2").value).toBe(42);
    expect(dst.getCell("A2").numFmt).toBe("0.00");
    expect(dst.getCell("B2").value).toBe(true);
    expect(dst.getCell("A3").value).toBeInstanceOf(Date);
    expect((dst.getCell("A3").value as Date).getTime()).toBe(
      new Date("2026-01-15T00:00:00.000Z").getTime(),
    );
    expect(dst.getCell("B3").value).toEqual({ error: "#REF!" });
    expect(dst.getCell("A4").formula).toBe("A2+1");
    expect(dst.getCell("A4").result).toBe(43);
    expect(dst.getColumn(1).width).toBe(20);
    expect(dst.getCell("A5").value).toBe("Merged");
    expect(dst.getCell("B5").isMerged).toBe(true);
  });

  it("reproduces string values through the inline-string path", () => {
    const dst = readBack(false);
    expect(dst.getCell("A1").value).toBe("Hello");
    expect(dst.getCell("B1").value).toBe("Hello");
    expect(dst.getCell("A2").value).toBe(42);
    expect(dst.getCell("B2").value).toBe(true);
    expect(dst.getCell("A5").value).toBe("Merged");
    expect(dst.getCell("B5").isMerged).toBe(true);
  });
});

// --- <hyperlinks> reconstruction ----------------------------------------------------------------------

describe("parseWorksheet — hyperlinks", () => {
  it("reconstructs { text, hyperlink, tooltip } from a <hyperlinks> block + rels map", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>` +
          `<hyperlinks><hyperlink ref="A1" r:id="rId1" tooltip="Visit"/></hyperlinks>`,
      ),
      ctx({
        sharedStrings: ["Anthropic"],
        rels: new Map([
          ["rId1", { type: "hyperlink", target: "https://anthropic.com", mode: "External" }],
        ]),
      }),
    );

    expect(s.getCell("A1").type).toBe(CellType.Hyperlink);
    expect(s.getCell("A1").value).toEqual({
      text: "Anthropic",
      hyperlink: "https://anthropic.com",
      tooltip: "Visit",
    });
  });

  it("omits tooltip from the model value when the XML carries none", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>` +
          `<hyperlinks><hyperlink ref="A1" r:id="rId1"/></hyperlinks>`,
      ),
      ctx({
        sharedStrings: ["Mail us"],
        rels: new Map([
          ["rId1", { type: "hyperlink", target: "mailto:hi@example.com", mode: "External" }],
        ]),
      }),
    );

    const value = s.getCell("A1").value;
    expect(value).toEqual({ text: "Mail us", hyperlink: "mailto:hi@example.com" });
    expect(value).not.toHaveProperty("tooltip");
  });

  it("applies a range ref to its top-left cell only", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"/></row></sheetData>` +
          `<hyperlinks><hyperlink ref="A1:B1" r:id="rId1"/></hyperlinks>`,
      ),
      ctx({
        sharedStrings: ["Link"],
        rels: new Map([
          ["rId1", { type: "hyperlink", target: "https://example.com", mode: "External" }],
        ]),
      }),
    );

    expect(s.getCell("A1").type).toBe(CellType.Hyperlink);
    expect(s.getCell("B1").type).not.toBe(CellType.Hyperlink);
  });

  it("leaves the cell as a plain string when rels are absent from the read context", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>` +
          `<hyperlinks><hyperlink ref="A1" r:id="rId1"/></hyperlinks>`,
      ),
      ctx({ sharedStrings: ["Plain"] }),
    );

    expect(s.getCell("A1").value).toBe("Plain");
    expect(s.getCell("A1").type).toBe(CellType.String);
  });

  it("skips a hyperlink whose r:id is not in the rels map (malformed package)", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>` +
          `<hyperlinks><hyperlink ref="A1" r:id="rId9"/></hyperlinks>`,
      ),
      ctx({ sharedStrings: ["Plain"], rels: new Map() }),
    );

    expect(s.getCell("A1").value).toBe("Plain");
  });
});

// --- comments (docs/COMMENTS-SPEC.md §2): <legacyDrawing>/rels -> comments{n}.xml -> cell.note --------

describe("parseWorksheet — comments", () => {
  const commentsXml =
    '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    "<authors><author>Author</author></authors>" +
    '<commentList><comment ref="B2" authorId="0"><text><t>hi</t></text></comment></commentList>' +
    "</comments>";

  it("follows <legacyDrawing>/rels to comments{n}.xml and applies cell.note", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(`<legacyDrawing r:id="rId2"/>`),
      ctx({
        rels: new Map([
          [
            "rId2",
            { type: REL_TYPE.vmlDrawing, target: "xl/drawings/vmlDrawing1.vml", mode: "Internal" },
          ],
          ["rId1", { type: REL_TYPE.comments, target: "xl/comments1.xml", mode: "Internal" }],
        ]),
        readPart: (path) => (path === "xl/comments1.xml" ? commentsXml : undefined),
      }),
    );
    expect(s.getCell("B2").note).toEqual({ texts: "hi", author: "Author" });
  });

  it("leaves every cell without a note when there is no <legacyDrawing>", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(""),
      ctx({
        rels: new Map([
          ["rId1", { type: REL_TYPE.comments, target: "xl/comments1.xml", mode: "Internal" }],
        ]),
        readPart: (path) => (path === "xl/comments1.xml" ? commentsXml : undefined),
      }),
    );
    expect(s.getCell("B2").note).toBeUndefined();
  });

  it("leaves every cell without a note when the legacyDrawing rId is absent from rels", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(`<legacyDrawing r:id="rId9"/>`),
      ctx({
        rels: new Map([
          ["rId1", { type: REL_TYPE.comments, target: "xl/comments1.xml", mode: "Internal" }],
        ]),
        readPart: (path) => (path === "xl/comments1.xml" ? commentsXml : undefined),
      }),
    );
    expect(s.getCell("B2").note).toBeUndefined();
  });

  it("leaves every cell without a note when the sheet has no comments-typed rel", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(`<legacyDrawing r:id="rId2"/>`),
      ctx({
        rels: new Map([
          [
            "rId2",
            { type: REL_TYPE.vmlDrawing, target: "xl/drawings/vmlDrawing1.vml", mode: "Internal" },
          ],
        ]),
        readPart: (path) => (path === "xl/comments1.xml" ? commentsXml : undefined),
      }),
    );
    expect(s.getCell("B2").note).toBeUndefined();
  });

  it("leaves every cell without a note when the comments part is missing from the package", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(`<legacyDrawing r:id="rId2"/>`),
      ctx({
        rels: new Map([
          [
            "rId2",
            { type: REL_TYPE.vmlDrawing, target: "xl/drawings/vmlDrawing1.vml", mode: "Internal" },
          ],
          ["rId1", { type: REL_TYPE.comments, target: "xl/comments1.xml", mode: "Internal" }],
        ]),
        readPart: () => undefined,
      }),
    );
    expect(s.getCell("B2").note).toBeUndefined();
  });

  it("leaves every cell without a note when readPart is absent from the context", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(`<legacyDrawing r:id="rId2"/>`),
      ctx({
        rels: new Map([
          [
            "rId2",
            { type: REL_TYPE.vmlDrawing, target: "xl/drawings/vmlDrawing1.vml", mode: "Internal" },
          ],
          ["rId1", { type: REL_TYPE.comments, target: "xl/comments1.xml", mode: "Internal" }],
        ]),
      }),
    );
    expect(s.getCell("B2").note).toBeUndefined();
  });

  it("skips a comment with a malformed ref, keeping any other well-formed comments", () => {
    const malformed =
      '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      "<authors><author>Author</author></authors>" +
      "<commentList>" +
      '<comment ref="not a ref!" authorId="0"><text><t>bad</t></text></comment>' +
      '<comment ref="C5" authorId="0"><text><t>good</t></text></comment>' +
      "</commentList>" +
      "</comments>";
    const s = sheet();
    parseWorksheet(
      s,
      ws(`<legacyDrawing r:id="rId2"/>`),
      ctx({
        rels: new Map([
          [
            "rId2",
            { type: REL_TYPE.vmlDrawing, target: "xl/drawings/vmlDrawing1.vml", mode: "Internal" },
          ],
          ["rId1", { type: REL_TYPE.comments, target: "xl/comments1.xml", mode: "Internal" }],
        ]),
        readPart: (path) => (path === "xl/comments1.xml" ? malformed : undefined),
      }),
    );
    expect(s.getCell("C5").note).toEqual({ texts: "good", author: "Author" });
  });

  it("inverts a rich-text comment into RichText texts", () => {
    const rich =
      '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      "<authors><author>Alice</author></authors>" +
      '<commentList><comment ref="C3" authorId="0"><text>' +
      "<r><rPr><b/></rPr><t>bold</t></r><r><t>plain</t></r>" +
      "</text></comment></commentList>" +
      "</comments>";
    const s = sheet();
    parseWorksheet(
      s,
      ws(`<legacyDrawing r:id="rId2"/>`),
      ctx({
        rels: new Map([
          [
            "rId2",
            { type: REL_TYPE.vmlDrawing, target: "xl/drawings/vmlDrawing1.vml", mode: "Internal" },
          ],
          ["rId1", { type: REL_TYPE.comments, target: "xl/comments1.xml", mode: "Internal" }],
        ]),
        readPart: (path) => (path === "xl/comments1.xml" ? rich : undefined),
      }),
    );
    expect(s.getCell("C3").note).toEqual({
      texts: { richText: [{ text: "bold", font: { bold: true } }, { text: "plain" }] },
      author: "Alice",
    });
  });
});

// --- <conditionalFormatting> reconstruction (CONDITIONAL-FORMATTING-SPEC §3) --------------------------

describe("parseWorksheet — conditional formatting", () => {
  it("inverts cellIs with one formula (golden)", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<conditionalFormatting sqref="A1:A10">` +
          `<cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan">` +
          `<formula>5</formula></cfRule></conditionalFormatting>`,
      ),
      ctx(),
    );
    expect(s.conditionalFormattings).toEqual([
      {
        ref: "A1:A10",
        rules: [{ type: "cellIs", operator: "greaterThan", formulae: [5], priority: 1 }],
      },
    ]);
  });

  it("inverts cellIs between into two formulae (golden)", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<conditionalFormatting sqref="B1:B5">` +
          `<cfRule type="cellIs" dxfId="0" priority="3" operator="between">` +
          `<formula>1</formula><formula>10</formula></cfRule></conditionalFormatting>`,
      ),
      ctx(),
    );
    expect(s.conditionalFormattings[0]?.rules).toEqual([
      { type: "cellIs", operator: "between", formulae: [1, 10], priority: 3 },
    ]);
  });

  it("keeps a non-numeric cellIs formula as a string (a cell ref or quoted text)", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<conditionalFormatting sqref="A1"><cfRule type="cellIs" dxfId="0" priority="1" operator="equal">` +
          `<formula>$B$1</formula></cfRule></conditionalFormatting>`,
      ),
      ctx(),
    );
    expect(s.conditionalFormattings[0]?.rules[0]).toMatchObject({ formulae: ["$B$1"] });
  });

  it("inverts expression (golden)", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<conditionalFormatting sqref="A1:A3">` +
          `<cfRule type="expression" dxfId="0" priority="2"><formula>$A1&gt;0</formula></cfRule>` +
          `</conditionalFormatting>`,
      ),
      ctx(),
    );
    expect(s.conditionalFormattings[0]?.rules).toEqual([
      { type: "expression", formulae: ["$A1>0"], priority: 2 },
    ]);
  });

  it("inverts containsText: reads text off the attribute, discards the derived formula (golden)", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<conditionalFormatting sqref="C2:C20">` +
          `<cfRule type="containsText" dxfId="0" priority="4" operator="containsText" text="foo">` +
          `<formula>NOT(ISERROR(SEARCH("foo",C2)))</formula></cfRule></conditionalFormatting>`,
      ),
      ctx(),
    );
    expect(s.conditionalFormattings[0]?.rules).toEqual([
      { type: "containsText", text: "foo", priority: 4 },
    ]);
  });

  it("inverts notContainsText / beginsWith / endsWith by their type attribute (golden)", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<conditionalFormatting sqref="A1">` +
          `<cfRule type="notContainsText" dxfId="0" priority="1" operator="notContains" text="bar">` +
          `<formula>ISERROR(SEARCH("bar",A1))</formula></cfRule>` +
          `<cfRule type="beginsWith" dxfId="0" priority="2" operator="beginsWith" text="Q">` +
          `<formula>LEFT(A1,LEN("Q"))="Q"</formula></cfRule>` +
          `<cfRule type="endsWith" dxfId="0" priority="3" operator="endsWith" text="Z">` +
          `<formula>RIGHT(A1,LEN("Z"))="Z"</formula></cfRule>` +
          `</conditionalFormatting>`,
      ),
      ctx(),
    );
    expect(s.conditionalFormattings[0]?.rules).toEqual([
      { type: "notContainsText", text: "bar", priority: 1 },
      { type: "beginsWith", text: "Q", priority: 2 },
      { type: "endsWith", text: "Z", priority: 3 },
    ]);
  });

  it("inverts top10, with percent/bottom only when the attribute is present (golden)", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<conditionalFormatting sqref="A1:A100">` +
          `<cfRule type="top10" dxfId="0" priority="4" percent="1" bottom="1" rank="10"/>` +
          `</conditionalFormatting>`,
      ),
      ctx(),
    );
    expect(s.conditionalFormattings[0]?.rules).toEqual([
      { type: "top10", rank: 10, percent: true, bottom: true, priority: 4 },
    ]);
  });

  it("top10 omits percent/bottom when the attributes are absent", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<conditionalFormatting sqref="A1:A100">` +
          `<cfRule type="top10" dxfId="0" priority="1" rank="5"/></conditionalFormatting>`,
      ),
      ctx(),
    );
    const rule = s.conditionalFormattings[0]?.rules[0];
    expect(rule).toEqual({ type: "top10", rank: 5, priority: 1 });
    expect(rule).not.toHaveProperty("percent");
    expect(rule).not.toHaveProperty("bottom");
  });

  it("inverts aboveAverage / belowAverage via the aboveAverage=0 flag (golden)", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<conditionalFormatting sqref="A1:A9">` +
          `<cfRule type="aboveAverage" dxfId="0" priority="5"/>` +
          `<cfRule type="aboveAverage" dxfId="0" priority="6" aboveAverage="0"/>` +
          `</conditionalFormatting>`,
      ),
      ctx(),
    );
    expect(s.conditionalFormattings[0]?.rules).toEqual([
      { type: "aboveAverage", priority: 5 },
      { type: "belowAverage", priority: 6 },
    ]);
  });

  it("inverts duplicateValues / uniqueValues (golden)", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<conditionalFormatting sqref="A1:A9">` +
          `<cfRule type="duplicateValues" dxfId="0" priority="6"/>` +
          `<cfRule type="uniqueValues" dxfId="0" priority="7"/>` +
          `</conditionalFormatting>`,
      ),
      ctx(),
    );
    expect(s.conditionalFormattings[0]?.rules).toEqual([
      { type: "duplicateValues", priority: 6 },
      { type: "uniqueValues", priority: 7 },
    ]);
  });

  it("inverts colorScale with no dxfId (golden)", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<conditionalFormatting sqref="C1:C20"><cfRule type="colorScale" priority="7"><colorScale>` +
          `<cfvo type="min"/><cfvo type="max"/>` +
          `<color rgb="FFF8696B"/><color rgb="FF63BE7B"/>` +
          `</colorScale></cfRule></conditionalFormatting>`,
      ),
      ctx(),
    );
    expect(s.conditionalFormattings[0]?.rules).toEqual([
      {
        type: "colorScale",
        cfvo: [{ type: "min" }, { type: "max" }],
        color: ["FFF8696B", "FF63BE7B"],
        priority: 7,
      },
    ]);
  });

  it("inverts dataBar (golden)", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<conditionalFormatting sqref="D1:D20"><cfRule type="dataBar" priority="8"><dataBar>` +
          `<cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/>` +
          `</dataBar></cfRule></conditionalFormatting>`,
      ),
      ctx(),
    );
    expect(s.conditionalFormattings[0]?.rules).toEqual([
      {
        type: "dataBar",
        cfvo: [{ type: "min" }, { type: "max" }],
        color: "FF638EC6",
        priority: 8,
      },
    ]);
  });

  it("inverts iconSet including percent cfvo vals (golden)", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<conditionalFormatting sqref="E1:E20">` +
          `<cfRule type="iconSet" priority="9"><iconSet iconSet="3TrafficLights1">` +
          `<cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/>` +
          `</iconSet></cfRule></conditionalFormatting>`,
      ),
      ctx(),
    );
    expect(s.conditionalFormattings[0]?.rules).toEqual([
      {
        type: "iconSet",
        iconSet: "3TrafficLights1",
        cfvo: [
          { type: "percent", value: 0 },
          { type: "percent", value: 33 },
          { type: "percent", value: 67 },
        ],
        priority: 9,
      },
    ]);
  });

  it("resolves dxfId into rule.style, including a bgColor fill highlight", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<conditionalFormatting sqref="A1:A10">` +
          `<cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan">` +
          `<formula>5</formula></cfRule></conditionalFormatting>`,
      ),
      ctx({
        styles: stubStyles({
          dxfs: {
            0: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFF8696B" } } },
          },
        }),
      }),
    );
    expect(s.conditionalFormattings[0]?.rules[0]).toEqual({
      type: "cellIs",
      operator: "greaterThan",
      formulae: [5],
      priority: 1,
      style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFF8696B" } } },
    });
  });

  it("leaves style unset when dxfId resolves to an empty differential style", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<conditionalFormatting sqref="A1:A10">` +
          `<cfRule type="duplicateValues" dxfId="0" priority="1"/></conditionalFormatting>`,
      ),
      ctx(),
    );
    expect(s.conditionalFormattings[0]?.rules[0]).not.toHaveProperty("style");
  });

  it("registers several sibling <conditionalFormatting> blocks (no outer wrapper)", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<conditionalFormatting sqref="A1:A9"><cfRule type="duplicateValues" dxfId="0" priority="1"/></conditionalFormatting>` +
          `<conditionalFormatting sqref="C1:C9"><cfRule type="uniqueValues" dxfId="0" priority="2"/></conditionalFormatting>`,
      ),
      ctx(),
    );
    expect(s.conditionalFormattings.map((cf) => cf.ref)).toEqual(["A1:A9", "C1:C9"]);
  });

  it("skips a block missing sqref, never throwing", () => {
    const s = sheet();
    expect(() =>
      parseWorksheet(
        s,
        ws(
          `<conditionalFormatting><cfRule type="duplicateValues" dxfId="0" priority="1"/></conditionalFormatting>`,
        ),
        ctx(),
      ),
    ).not.toThrow();
    expect(s.conditionalFormattings).toEqual([]);
  });

  it("skips an individual cfRule with an unrecognized type, keeping its siblings", () => {
    const s = sheet();
    expect(() =>
      parseWorksheet(
        s,
        ws(
          `<conditionalFormatting sqref="A1:A9">` +
            `<cfRule type="timePeriod" dxfId="0" priority="1"/>` +
            `<cfRule type="duplicateValues" dxfId="0" priority="2"/>` +
            `</conditionalFormatting>`,
        ),
        ctx(),
      ),
    ).not.toThrow();
    expect(s.conditionalFormattings[0]?.rules).toEqual([{ type: "duplicateValues", priority: 2 }]);
  });

  it("throws InvalidConditionalFormattingError for a non-numeric priority, never a silent NaN", () => {
    const s = sheet();
    expect(() =>
      parseWorksheet(
        s,
        ws(
          `<conditionalFormatting sqref="A1:A9">` +
            `<cfRule type="duplicateValues" dxfId="0" priority="x"/></conditionalFormatting>`,
        ),
        ctx(),
      ),
    ).toThrow(InvalidConditionalFormattingError);
  });

  it("throws InvalidConditionalFormattingError for a non-numeric rank, never a silent NaN", () => {
    const s = sheet();
    expect(() =>
      parseWorksheet(
        s,
        ws(
          `<conditionalFormatting sqref="A1:A9">` +
            `<cfRule type="top10" dxfId="0" priority="1" rank="x"/></conditionalFormatting>`,
        ),
        ctx(),
      ),
    ).toThrow(InvalidConditionalFormattingError);
  });

  it("skips the whole sheet's conditional formatting when the sqref fails Range validation", () => {
    const s = sheet();
    expect(() =>
      parseWorksheet(
        s,
        ws(
          `<conditionalFormatting sqref="@@@">` +
            `<cfRule type="duplicateValues" dxfId="0" priority="1"/></conditionalFormatting>`,
        ),
        ctx(),
      ),
    ).not.toThrow();
    expect(s.conditionalFormattings).toEqual([]);
  });
});

// --- <dataValidations> reconstruction ------------------------------------------------------------

describe("parseWorksheet — data validations", () => {
  it("parses an inline list dropdown", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<dataValidations count="1"><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1">` +
          `<formula1>"Yes,No,Maybe"</formula1></dataValidation></dataValidations>`,
      ),
      ctx(),
    );

    expect(s.dataValidations.find("A1")).toEqual({
      type: "list",
      allowBlank: true,
      showInputMessage: true,
      showErrorMessage: true,
      formulae: ['"Yes,No,Maybe"'],
    });
  });

  it("parses a same-sheet range list dropdown", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<dataValidations count="1"><dataValidation type="list" sqref="A1:A10">` +
          `<formula1>$E$1:$E$5</formula1></dataValidation></dataValidations>`,
      ),
      ctx(),
    );

    expect(s.dataValidations.find("A5")).toEqual({ type: "list", formulae: ["$E$1:$E$5"] });
  });

  it("parses a decimal between with two numeric formulae", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<dataValidations count="1"><dataValidation type="decimal" operator="between" sqref="B2">` +
          `<formula1>0.5</formula1><formula2>9.5</formula2></dataValidation></dataValidations>`,
      ),
      ctx(),
    );

    expect(s.dataValidations.find("B2")).toEqual({
      type: "decimal",
      operator: "between",
      formulae: [0.5, 9.5],
    });
  });

  it('defaults operator to "between" when absent on an operator type (ECMA-376/ExcelJS default)', () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<dataValidations count="1"><dataValidation type="decimal" sqref="B2">` +
          `<formula1>0</formula1><formula2>100</formula2></dataValidation></dataValidations>`,
      ),
      ctx(),
    );

    expect(s.dataValidations.find("B2")).toEqual({
      type: "decimal",
      operator: "between",
      formulae: [0, 100],
    });
  });

  it("keeps a non-numeric formula (cell ref) as a string for an operator type", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<dataValidations count="1"><dataValidation type="whole" operator="greaterThan" sqref="C3">` +
          `<formula1>$D$1</formula1></dataValidation></dataValidations>`,
      ),
      ctx(),
    );

    expect(s.dataValidations.find("C3")).toEqual({
      type: "whole",
      operator: "greaterThan",
      formulae: ["$D$1"],
    });
  });

  it("parses a custom rule with no operator", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<dataValidations count="1"><dataValidation type="custom" sqref="D4">` +
          `<formula1>ISNUMBER(D4)</formula1></dataValidation></dataValidations>`,
      ),
      ctx(),
    );

    expect(s.dataValidations.find("D4")).toEqual({
      type: "custom",
      formulae: ["ISNUMBER(D4)"],
    });
  });

  it("ignores a stray operator attribute on a list type (never gated to that type)", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<dataValidations count="1"><dataValidation type="list" operator="between" sqref="A1">` +
          `<formula1>"a,b"</formula1></dataValidation></dataValidations>`,
      ),
      ctx(),
    );

    expect(s.dataValidations.find("A1")?.operator).toBeUndefined();
  });

  it("parses the message/error attributes", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<dataValidations count="1"><dataValidation type="textLength" operator="lessThanOrEqual" showErrorMessage="1" errorStyle="warning" errorTitle="Too long" error="Keep it short" promptTitle="Note" prompt="Enter a short code" sqref="E5">` +
          `<formula1>8</formula1></dataValidation></dataValidations>`,
      ),
      ctx(),
    );

    expect(s.dataValidations.find("E5")).toEqual({
      type: "textLength",
      operator: "lessThanOrEqual",
      showErrorMessage: true,
      errorStyle: "warning",
      errorTitle: "Too long",
      error: "Keep it short",
      promptTitle: "Note",
      prompt: "Enter a short code",
      formulae: [8],
    });
  });

  it("inverts showDropDown: attr present → model false", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<dataValidations count="1"><dataValidation type="list" showDropDown="1" sqref="F6">` +
          `<formula1>"a,b"</formula1></dataValidation></dataValidations>`,
      ),
      ctx(),
    );

    expect(s.dataValidations.find("F6")?.showDropDown).toBe(false);
  });

  it("leaves showDropDown unset when the attribute is absent (dropdown shown)", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<dataValidations count="1"><dataValidation type="list" sqref="F7">` +
          `<formula1>"a,b"</formula1></dataValidation></dataValidations>`,
      ),
      ctx(),
    );

    expect(s.dataValidations.find("F7")).not.toHaveProperty("showDropDown");
  });

  it("registers each entry under its own sqref, independent of document order relative to other blocks", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<dataValidations count="2">` +
          `<dataValidation type="list" sqref="A1"><formula1>"a,b"</formula1></dataValidation>` +
          `<dataValidation type="custom" sqref="B1"><formula1>TRUE()</formula1></dataValidation>` +
          `</dataValidations>`,
      ),
      ctx(),
    );

    expect(s.dataValidations.find("A1")?.type).toBe("list");
    expect(s.dataValidations.find("B1")?.type).toBe("custom");
  });

  it("resolves an address against a multi-range sqref", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<dataValidations count="1"><dataValidation type="list" sqref="A1:A3 C1:C3">` +
          `<formula1>"a,b"</formula1></dataValidation></dataValidations>`,
      ),
      ctx(),
    );

    expect(s.dataValidations.find("C2")?.type).toBe("list");
    expect(s.dataValidations.find("B2")).toBeUndefined();
  });

  it("skips an entry missing sqref, never throwing", () => {
    const s = sheet();
    expect(() =>
      parseWorksheet(
        s,
        ws(
          `<dataValidations count="1"><dataValidation type="list">` +
            `<formula1>"a,b"</formula1></dataValidation></dataValidations>`,
        ),
        ctx(),
      ),
    ).not.toThrow();
    expect(s.dataValidations.size).toBe(0);
  });

  it("skips an entry with an unrecognized type, never throwing", () => {
    const s = sheet();
    expect(() =>
      parseWorksheet(
        s,
        ws(
          `<dataValidations count="1"><dataValidation type="bogus" sqref="A1">` +
            `<formula1>x</formula1></dataValidation></dataValidations>`,
        ),
        ctx(),
      ),
    ).not.toThrow();
    expect(s.dataValidations.find("A1")).toBeUndefined();
  });
});

// --- <sheetViews>/<pane> reconstruction (PANES-AUTOFILTER-SPEC §3) ------------------------------------

describe("parseWorksheet — sheetViews / panes", () => {
  it("leaves views empty for a bare <sheetView> (default, golden)", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(`<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData/>`),
      ctx(),
    );
    expect(s.views).toEqual([]);
  });

  it("parses a frozen pane (golden)", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<sheetViews><sheetView workbookViewId="0">` +
          `<pane xSplit="1" ySplit="1" topLeftCell="B2" state="frozen"/>` +
          `</sheetView></sheetViews><sheetData/>`,
      ),
      ctx(),
    );
    expect(s.views).toEqual([{ state: "frozen", xSplit: 1, ySplit: 1, topLeftCell: "B2" }]);
  });

  it("parses a split pane with activePane (golden)", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<sheetViews><sheetView workbookViewId="0">` +
          `<pane xSplit="2000" ySplit="3000" topLeftCell="C4" activePane="bottomRight" state="split"/>` +
          `</sheetView></sheetViews><sheetData/>`,
      ),
      ctx(),
    );
    expect(s.views).toEqual([
      { state: "split", xSplit: 2000, ySplit: 3000, topLeftCell: "C4", activePane: "bottomRight" },
    ]);
  });

  it("reads showGridLines=false off the sheetView even with no pane", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(`<sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews><sheetData/>`),
      ctx(),
    );
    expect(s.views).toEqual([{ showGridLines: false }]);
  });

  it("throws InvalidPaneError for a non-numeric xSplit, never a silent NaN", () => {
    const s = sheet();
    expect(() =>
      parseWorksheet(
        s,
        ws(
          `<sheetViews><sheetView workbookViewId="0">` +
            `<pane state="frozen" xSplit="abc"/>` +
            `</sheetView></sheetViews><sheetData/>`,
        ),
        ctx(),
      ),
    ).toThrow(InvalidPaneError);
  });

  it("parses zoomScale off the sheetView", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(`<sheetViews><sheetView workbookViewId="0" zoomScale="150"/></sheetViews><sheetData/>`),
      ctx(),
    );
    expect(s.views).toEqual([{ zoomScale: 150 }]);
  });

  it("drops zoomScale=100 (the default) rather than materializing a view", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(`<sheetViews><sheetView workbookViewId="0" zoomScale="100"/></sheetViews><sheetData/>`),
      ctx(),
    );
    expect(s.views).toEqual([]);
  });

  it("parses zoomScale alongside a frozen pane", () => {
    const s = sheet();
    parseWorksheet(
      s,
      ws(
        `<sheetViews><sheetView workbookViewId="0" zoomScale="75">` +
          `<pane xSplit="1" ySplit="1" topLeftCell="B2" state="frozen"/>` +
          `</sheetView></sheetViews><sheetData/>`,
      ),
      ctx(),
    );
    expect(s.views).toEqual([
      { zoomScale: 75, state: "frozen", xSplit: 1, ySplit: 1, topLeftCell: "B2" },
    ]);
  });
});

// --- <autoFilter> reconstruction (PANES-AUTOFILTER-SPEC §3) -------------------------------------------

describe("parseWorksheet — autoFilter", () => {
  it("parses <autoFilter ref> (golden)", () => {
    const s = sheet();
    parseWorksheet(s, ws(`<sheetData/><autoFilter ref="A1:D1"/>`), ctx());
    expect(s.autoFilter).toBe("A1:D1");
  });

  it("tolerates filterColumn/sortState children without choking (out of scope)", () => {
    const s = sheet();
    expect(() =>
      parseWorksheet(
        s,
        ws(
          `<sheetData/><autoFilter ref="A1:C1">` +
            `<filterColumn colId="0"><filters><filter val="x"/></filters></filterColumn>` +
            `<sortState ref="A1:C1"/></autoFilter>`,
        ),
        ctx(),
      ),
    ).not.toThrow();
    expect(s.autoFilter).toBe("A1:C1");
  });
});
