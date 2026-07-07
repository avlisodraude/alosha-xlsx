import { describe, expect, it } from "vitest";
import { Workbook } from "../../../src/model/index.js";
import { REL_TYPE, type WorksheetRel } from "../../../src/xlsx/parts/constants.js";
import { renderWorksheet } from "../../../src/xlsx/parts/worksheet.js";
import { renderWorksheetRels } from "../../../src/xlsx/parts/worksheet-rels.js";
import { createSharedStrings } from "../../../src/xlsx/shared-strings.js";
import { createStyles } from "../../../src/xlsx/styles.js";

/** Render a single-sheet worksheet through the real pools, returning its xml + rels. */
function render(build: (ws: ReturnType<Workbook["addWorksheet"]>) => void) {
  const ws = new Workbook().addWorksheet("Data");
  build(ws);
  return renderWorksheet(ws, {
    styles: createStyles(),
    sharedStrings: createSharedStrings(),
    useSharedStrings: true,
    date1904: false,
  });
}

describe("worksheet writer — hyperlinks", () => {
  it("emits the display text as a string cell AND a <hyperlinks> entry with a matching r:id", () => {
    const { xml, rels } = render((ws) => {
      ws.getCell("A1").value = { text: "Anthropic", hyperlink: "https://anthropic.com" };
    });

    // The cell itself is an ordinary shared-string cell carrying the display text (index 0).
    expect(xml).toContain('<c r="A1" t="s"><v>0</v></c>');
    // A sheet-level hyperlink record points at the same cell via a worksheet-local rId.
    expect(xml).toContain('<hyperlinks><hyperlink ref="A1" r:id="rId1"/></hyperlinks>');
    // The <hyperlinks> block sits after any <mergeCells> and before the closing tag.
    expect(xml).toMatch(/<\/sheetData>.*<hyperlinks>.*<\/worksheet>$/s);

    // The rels doc resolves that same rId1 to the external target.
    expect(rels).toBeDefined();
    expect(rels).toContain(
      `<Relationship Id="rId1" Type="${REL_TYPE.hyperlink}" Target="https://anthropic.com" TargetMode="External"/>`,
    );
  });

  it("includes the tooltip attribute only when present", () => {
    const withTip = render((ws) => {
      ws.getCell("A1").value = {
        text: "Mail us",
        hyperlink: "mailto:hi@example.com",
        tooltip: "Send an email",
      };
    });
    expect(withTip.xml).toContain('<hyperlink ref="A1" r:id="rId1" tooltip="Send an email"/>');

    const withoutTip = render((ws) => {
      ws.getCell("A1").value = { text: "Plain", hyperlink: "https://example.com" };
    });
    expect(withoutTip.xml).toContain('<hyperlink ref="A1" r:id="rId1"/>');
    expect(withoutTip.xml).not.toContain("tooltip=");
  });

  it("assigns independent local rIds per hyperlink in document order", () => {
    const { xml, rels } = render((ws) => {
      ws.getCell("A1").value = { text: "One", hyperlink: "https://one.example" };
      ws.getCell("A2").value = { text: "Two", hyperlink: "mailto:two@example.com" };
    });
    expect(xml).toContain('<hyperlink ref="A1" r:id="rId1"/>');
    expect(xml).toContain('<hyperlink ref="A2" r:id="rId2"/>');
    expect(rels).toContain('Id="rId1" ');
    expect(rels).toContain('Id="rId2" ');
    expect(rels).toContain('Target="https://one.example"');
    expect(rels).toContain('Target="mailto:two@example.com"');
  });

  it("produces no <hyperlinks> block and no rels when the sheet has none", () => {
    const { xml, rels } = render((ws) => {
      ws.getCell("A1").value = "just text";
    });
    expect(xml).not.toContain("<hyperlinks>");
    expect(rels).toBeUndefined();
  });
});

describe("renderWorksheetRels", () => {
  it("renders external hyperlink relationships (golden)", () => {
    const rels: WorksheetRel[] = [
      { rId: "rId1", target: "https://example.com", mode: "External" },
      { rId: "rId2", target: "mailto:a@b.c", mode: "External" },
    ];
    expect(renderWorksheetRels(rels)).toBe(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        `<Relationship Id="rId1" Type="${REL_TYPE.hyperlink}" Target="https://example.com" TargetMode="External"/>` +
        `<Relationship Id="rId2" Type="${REL_TYPE.hyperlink}" Target="mailto:a@b.c" TargetMode="External"/>` +
        "</Relationships>",
    );
  });

  it("escapes ampersands in the target URL", () => {
    const rels: WorksheetRel[] = [
      { rId: "rId1", target: "https://example.com/?a=1&b=2", mode: "External" },
    ];
    expect(renderWorksheetRels(rels)).toContain("https://example.com/?a=1&amp;b=2");
  });
});
