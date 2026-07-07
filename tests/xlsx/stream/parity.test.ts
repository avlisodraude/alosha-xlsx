// Track E2 parity matrix: a fixture exercising every feature the buffered writer supports beyond bare
// values — cell styles, a merge, frozen panes, an auto-filter, conditional formatting, a data
// validation, cell comments, a hyperlink, and an embedded image — applied identically to a buffered
// model workbook and to the streaming WorkbookWriter. Both archives are parsed back with the reader and
// deep-diffed: the streamed output must reconstruct to the same model as writeWorkbookBuffer's. Positive
// assertions on the streamed model then confirm each feature is actually present (a deep-diff alone
// would pass even if the reader silently dropped a feature from both sides).

import { describe, expect, it } from "vitest";
import type { CellValue, Style, Workbook } from "../../../src/model/index.js";
import { Workbook as WorkbookModel } from "../../../src/model/index.js";
import { readWorkbookBuffer } from "../../../src/xlsx/read-workbook.js";
import { WorkbookWriter } from "../../../src/xlsx/stream/index.js";
import { writeWorkbookBuffer } from "../../../src/xlsx/write-workbook.js";

// Fixed doc metadata so the streamed and buffered docProps are identical between two workbooks.
const CREATED = new Date("2026-01-02T03:04:05.000Z");
const MODIFIED = new Date("2026-02-03T04:05:06.000Z");
const CREATOR = "Streaming E2";

// A tiny valid-enough PNG header — the writer stores the bytes verbatim and the reader hands them back.
const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

const BOLD: Partial<Style> = { font: { bold: true } };
const FILL: Partial<Style> = {
  fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF2CC" } },
};

/** A per-cell decoration applied after the row's values are set (identically on both paths). */
interface Decor {
  col: number;
  style?: Partial<Style>;
  numFmt?: string;
  note?: { texts: string; author?: string };
  dataValidation?: { type: "list"; formulae: string[]; allowBlank?: boolean };
}

interface RowSpec {
  values: CellValue[];
  decor?: Decor[];
}

// Header + four data rows. Row 1 is a styled/commented header; B-column carries numbers (numFmt +
// conditional formatting); C-column carries a validated list; one hyperlink; a merge on the last row.
const ROWS: RowSpec[] = [
  {
    values: ["Name", "Score", "Active"],
    decor: [
      { col: 1, style: BOLD, note: { texts: "Row header", author: "Reviewer" } },
      { col: 2, style: BOLD },
      { col: 3, style: BOLD },
    ],
  },
  {
    values: ["alpha", 42, "Yes"],
    decor: [
      { col: 2, numFmt: "0.00" },
      { col: 3, dataValidation: { type: "list", formulae: ['"Yes,No"'], allowBlank: true } },
    ],
  },
  {
    values: [{ text: "bravo", hyperlink: "https://example.com", tooltip: "open" }, 7.5, "No"],
    decor: [{ col: 2, style: FILL }],
  },
  { values: ["charlie", { formula: "B2+B3", result: 49.5 }, "Yes"] },
  { values: ["Total", 99, ""] },
];

const AUTO_FILTER = "A1:C1";
const FROZEN_VIEW = { state: "frozen" as const, xSplit: 0, ySplit: 1, topLeftCell: "A2" };
const MERGE = "A5:C5";
const CF = {
  ref: "B2:B5",
  rules: [
    { type: "cellIs" as const, operator: "greaterThan" as const, formulae: [10], style: FILL },
  ],
};
const IMAGE_ANCHOR = "E2:F6";

/** Apply the fixture to a buffered model worksheet. */
function buildBuffered(): Workbook {
  const wb = new WorkbookModel();
  wb.creator = CREATOR;
  wb.created = CREATED;
  wb.modified = MODIFIED;

  const ws = wb.addWorksheet("Report");
  ws.views = [FROZEN_VIEW];
  ws.autoFilter = AUTO_FILTER;

  ROWS.forEach((spec, i) => {
    const row = ws.getRow(i + 1);
    row.values = spec.values;
    for (const d of spec.decor ?? []) {
      const cell = row.getCell(d.col);
      if (d.style) cell.style = { ...cell.style, ...d.style };
      if (d.numFmt) cell.numFmt = d.numFmt;
      if (d.note) cell.note = d.note;
      if (d.dataValidation) cell.dataValidation = d.dataValidation;
    }
  });

  ws.mergeCells(MERGE);
  ws.addConditionalFormatting(CF);
  const imageId = wb.addImage({ extension: "png", buffer: PNG });
  ws.addImage(imageId, IMAGE_ANCHOR);
  return wb;
}

/** Apply the identical fixture through the streaming writer, returning the whole archive. */
async function buildStreamed(): Promise<Uint8Array> {
  const wb = new WorkbookWriter({ useSharedStrings: false });
  wb.creator = CREATOR;
  wb.created = CREATED;
  wb.modified = MODIFIED;

  const ws = wb.addWorksheet("Report");
  // Panes + auto-filter are read by the prologue, so configure them before the first row commits.
  ws.views = [FROZEN_VIEW];
  ws.autoFilter = AUTO_FILTER;

  for (const spec of ROWS) {
    const row = ws.addRow(spec.values);
    for (const d of spec.decor ?? []) {
      const cell = row.getCell(d.col);
      if (d.style) cell.style = { ...cell.style, ...d.style };
      if (d.numFmt) cell.numFmt = d.numFmt;
      if (d.note) cell.note = d.note;
      if (d.dataValidation) cell.dataValidation = d.dataValidation;
    }
    row.commit();
  }

  ws.mergeCells(MERGE);
  ws.addConditionalFormatting(CF);
  const imageId = wb.addImage({ extension: "png", buffer: PNG });
  ws.addImage(imageId, IMAGE_ANCHOR);
  ws.commit();

  await wb.commit();
  return collect(wb);
}

async function collect(source: AsyncIterable<Uint8Array>): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  let total = 0;
  for await (const chunk of source) {
    chunks.push(chunk);
    total += chunk.length;
  }
  const out = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    out.set(chunk, offset);
    offset += chunk.length;
  }
  return out;
}

/** Stable JSON of a parsed workbook model (Dates → ISO), for structural deep-diffing. */
function modelJson(wb: Workbook): string {
  return JSON.stringify(wb.model, (_k, v) => (v instanceof Date ? v.toISOString() : v), 2);
}

describe("streaming parity — full feature matrix", () => {
  it("streams a workbook structurally equal to the buffered writer across every feature", async () => {
    const streamed = await buildStreamed();
    const buffered = writeWorkbookBuffer(buildBuffered(), { useSharedStrings: false });

    expect(modelJson(readWorkbookBuffer(streamed))).toEqual(
      modelJson(readWorkbookBuffer(buffered)),
    );
  });

  it("reconstructs each individual feature from the streamed archive", async () => {
    const wb = readWorkbookBuffer(await buildStreamed());
    const ws = wb.getWorksheet("Report");
    if (!ws) throw new Error("Report sheet missing");

    // Values + a formula + a hyperlink's display text.
    expect(ws.getCell("A2").value).toBe("alpha");
    expect(ws.getCell("B2").value).toBe(42);
    expect(ws.getCell("A4").value).toBe("charlie");
    expect(ws.getCell("B4").formula).toBe("B2+B3");
    expect(ws.getCell("B4").result).toBe(49.5);
    // A3 is a hyperlink cell — reconstructed as a link object carrying its display text + target.
    expect(ws.getCell("A3").value).toMatchObject({
      text: "bravo",
      hyperlink: "https://example.com",
    });

    // Cell style + number format flow through renderCell's `s` attr.
    expect(ws.getCell("A1").font?.bold).toBe(true);
    expect(ws.getCell("B2").numFmt).toBe("0.00");

    // Frozen panes (prologue), auto-filter + merge (epilogue).
    expect(ws.views[0]?.state).toBe("frozen");
    expect(ws.views[0]?.ySplit).toBe(1);
    expect(ws.autoFilter).toBe(AUTO_FILTER);
    expect(ws.getCell("B5").isMerged).toBe(true);

    // Conditional formatting.
    expect(ws.conditionalFormattings[0]?.ref).toBe("B2:B5");
    expect(ws.conditionalFormattings[0]?.rules[0]?.type).toBe("cellIs");

    // Data validation.
    expect(ws.getCell("C2").dataValidation?.type).toBe("list");

    // Comment (collected across the row commits, flushed at ws.commit()). The reader normalises the
    // note text to rich-text runs, so read the run text back.
    const noteTexts = ws.getCell("A1").note?.texts;
    const noteText =
      typeof noteTexts === "string"
        ? noteTexts
        : (noteTexts?.richText.map((r) => r.text).join("") ?? "");
    expect(noteText).toBe("Row header");

    // Embedded image: placement + byte-identical media recovered.
    expect(ws.images[0]?.imageId).toBe(0);
    expect(wb.getImage(0)).toEqual({ buffer: PNG, extension: "png" });
  });
});
