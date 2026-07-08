import { describe, expect, it } from "vitest";
import type { Style } from "../../../src/model/index.js";
import { Workbook } from "../../../src/model/index.js";
import { renderWorksheet } from "../../../src/xlsx/parts/worksheet.js";
import { createSharedStrings } from "../../../src/xlsx/shared-strings.js";
import { createStyles, type Styles } from "../../../src/xlsx/styles.js";

/** Render a single-sheet worksheet through the real pools; returns the sheet xml and the styles pool
 *  (so a test can also assert the `<dxfs>` it interned). */
function render(build: (ws: ReturnType<Workbook["addWorksheet"]>) => void): {
  xml: string;
  styles: Styles;
} {
  const ws = new Workbook().addWorksheet("CF");
  build(ws);
  const styles = createStyles();
  const xml = renderWorksheet(ws, {
    styles,
    sharedStrings: createSharedStrings(),
    useSharedStrings: true,
    date1904: false,
  }).xml;
  return { xml, styles };
}

const HIGHLIGHT: Partial<Style> = {
  fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFF8696B" } },
};

describe("worksheet writer — conditional formatting", () => {
  it("cellIs with one formula (golden)", () => {
    const { xml } = render((ws) => {
      ws.addConditionalFormatting({
        ref: "A1:A10",
        rules: [{ type: "cellIs", operator: "greaterThan", formulae: [5], priority: 1 }],
      });
    });
    expect(xml).toContain(
      '<conditionalFormatting sqref="A1:A10">' +
        '<cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan">' +
        "<formula>5</formula>" +
        "</cfRule>" +
        "</conditionalFormatting>",
    );
  });

  it("cellIs between emits two formulae (golden)", () => {
    const { xml } = render((ws) => {
      ws.addConditionalFormatting({
        ref: "B1:B5",
        rules: [{ type: "cellIs", operator: "between", formulae: [1, 10], priority: 3 }],
      });
    });
    expect(xml).toContain(
      '<cfRule type="cellIs" dxfId="0" priority="3" operator="between">' +
        "<formula>1</formula><formula>10</formula>" +
        "</cfRule>",
    );
  });

  it("expression (golden)", () => {
    const { xml } = render((ws) => {
      ws.addConditionalFormatting({
        ref: "A1:A3",
        rules: [{ type: "expression", formulae: ["$A1>0"], priority: 2 }],
      });
    });
    expect(xml).toContain(
      '<cfRule type="expression" dxfId="0" priority="2"><formula>$A1&gt;0</formula></cfRule>',
    );
  });

  it("containsText emits text attr + SEARCH formula anchored on the ref top-left (golden)", () => {
    const { xml } = render((ws) => {
      ws.addConditionalFormatting({
        ref: "C2:C20",
        rules: [{ type: "containsText", text: "foo", priority: 4 }],
      });
    });
    expect(xml).toContain(
      '<cfRule type="containsText" dxfId="0" priority="4" operator="containsText" text="foo">' +
        '<formula>NOT(ISERROR(SEARCH("foo",C2)))</formula>' +
        "</cfRule>",
    );
  });

  it("notContainsText uses operator notContains + ISERROR formula (golden)", () => {
    const { xml } = render((ws) => {
      ws.addConditionalFormatting({
        ref: "A1",
        rules: [{ type: "notContainsText", text: "bar", priority: 1 }],
      });
    });
    expect(xml).toContain(
      '<cfRule type="notContainsText" dxfId="0" priority="1" operator="notContains" text="bar">' +
        '<formula>ISERROR(SEARCH("bar",A1))</formula>' +
        "</cfRule>",
    );
  });

  it("beginsWith / endsWith formulae (golden)", () => {
    const begins = render((ws) => {
      ws.addConditionalFormatting({
        ref: "A1",
        rules: [{ type: "beginsWith", text: "Q", priority: 1 }],
      });
    }).xml;
    expect(begins).toContain('<formula>LEFT(A1,LEN("Q"))="Q"</formula>');

    const ends = render((ws) => {
      ws.addConditionalFormatting({
        ref: "A1",
        rules: [{ type: "endsWith", text: "Z", priority: 1 }],
      });
    }).xml;
    expect(ends).toContain('<formula>RIGHT(A1,LEN("Z"))="Z"</formula>');
  });

  it("escapes embedded quotes and commas in text-rule formulae without corrupting the formula", () => {
    const quote = render((ws) => {
      ws.addConditionalFormatting({
        ref: "A1:A5",
        rules: [{ type: "containsText", text: 'say "hi"', priority: 1 }],
      });
    }).xml;
    expect(quote).toContain('<formula>NOT(ISERROR(SEARCH("say ""hi""",A1)))</formula>');
    expect(quote).toContain('text="say &quot;hi&quot;"');

    const comma = render((ws) => {
      ws.addConditionalFormatting({
        ref: "A1:A5",
        rules: [{ type: "beginsWith", text: "a,b", priority: 1 }],
      });
    }).xml;
    expect(comma).toContain('<formula>LEFT(A1,LEN("a,b"))="a,b"</formula>');
  });

  it("top10 emits rank, and percent/bottom only when true (golden)", () => {
    const { xml } = render((ws) => {
      ws.addConditionalFormatting({
        ref: "A1:A100",
        rules: [{ type: "top10", rank: 10, percent: true, bottom: true, priority: 4 }],
      });
    });
    expect(xml).toContain(
      '<cfRule type="top10" dxfId="0" priority="4" percent="1" bottom="1" rank="10"/>',
    );
  });

  it("top10 omits percent/bottom when false", () => {
    const { xml } = render((ws) => {
      ws.addConditionalFormatting({
        ref: "A1:A100",
        rules: [{ type: "top10", rank: 5, priority: 1 }],
      });
    });
    expect(xml).toContain('<cfRule type="top10" dxfId="0" priority="1" rank="5"/>');
    expect(xml).not.toContain("percent=");
    expect(xml).not.toContain("bottom=");
  });

  it("aboveAverage / belowAverage (golden)", () => {
    const above = render((ws) => {
      ws.addConditionalFormatting({
        ref: "A1:A9",
        rules: [{ type: "aboveAverage", priority: 5 }],
      });
    }).xml;
    expect(above).toContain('<cfRule type="aboveAverage" dxfId="0" priority="5"/>');

    const below = render((ws) => {
      ws.addConditionalFormatting({
        ref: "A1:A9",
        rules: [{ type: "belowAverage", priority: 6 }],
      });
    }).xml;
    expect(below).toContain(
      '<cfRule type="aboveAverage" dxfId="0" priority="6" aboveAverage="0"/>',
    );
  });

  it("duplicateValues / uniqueValues (golden)", () => {
    const dup = render((ws) => {
      ws.addConditionalFormatting({
        ref: "A1:A9",
        rules: [{ type: "duplicateValues", priority: 6 }],
      });
    }).xml;
    expect(dup).toContain('<cfRule type="duplicateValues" dxfId="0" priority="6"/>');

    const uniq = render((ws) => {
      ws.addConditionalFormatting({
        ref: "A1:A9",
        rules: [{ type: "uniqueValues", priority: 7 }],
      });
    }).xml;
    expect(uniq).toContain('<cfRule type="uniqueValues" dxfId="0" priority="7"/>');
  });

  it("colorScale carries cfvo/color and no dxfId (golden)", () => {
    const { xml } = render((ws) => {
      ws.addConditionalFormatting({
        ref: "C1:C20",
        rules: [
          {
            type: "colorScale",
            cfvo: [{ type: "min" }, { type: "max" }],
            color: ["FFF8696B", "FF63BE7B"],
            priority: 7,
          },
        ],
      });
    });
    expect(xml).toContain(
      '<cfRule type="colorScale" priority="7"><colorScale>' +
        '<cfvo type="min"/><cfvo type="max"/>' +
        '<color rgb="FFF8696B"/><color rgb="FF63BE7B"/>' +
        "</colorScale></cfRule>",
    );
    expect(xml).not.toContain('type="colorScale" dxfId');
  });

  it("dataBar (golden)", () => {
    const { xml } = render((ws) => {
      ws.addConditionalFormatting({
        ref: "D1:D20",
        rules: [
          {
            type: "dataBar",
            cfvo: [{ type: "min" }, { type: "max" }],
            color: "FF638EC6",
            priority: 8,
          },
        ],
      });
    });
    expect(xml).toContain(
      '<cfRule type="dataBar" priority="8"><dataBar>' +
        '<cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/>' +
        "</dataBar></cfRule>",
    );
  });

  it("iconSet carries the preset + cfvo val for percent stops (golden)", () => {
    const { xml } = render((ws) => {
      ws.addConditionalFormatting({
        ref: "E1:E20",
        rules: [
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
        ],
      });
    });
    expect(xml).toContain(
      '<cfRule type="iconSet" priority="9"><iconSet iconSet="3TrafficLights1">' +
        '<cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/>' +
        "</iconSet></cfRule>",
    );
  });

  it("a dxf-bearing rule interns its style into the <dxfs> pool (bgColor highlight)", () => {
    const { xml, styles } = render((ws) => {
      ws.addConditionalFormatting({
        ref: "A1:A10",
        rules: [
          { type: "cellIs", operator: "greaterThan", formulae: [5], style: HIGHLIGHT, priority: 1 },
        ],
      });
    });
    // The rule references dxfId 0…
    expect(xml).toContain('type="cellIs" dxfId="0"');
    // …and styles.xml now carries exactly one dxf whose colour is in <bgColor> (not <fgColor>).
    const stylesXml = styles.toXml();
    expect(stylesXml).toContain(
      '<dxfs count="1"><dxf><fill><patternFill><bgColor rgb="FFF8696B"/></patternFill></fill></dxf></dxfs>',
    );
    expect(stylesXml).not.toContain("fgColor");
  });

  it("dedups identical dxfs and assigns distinct ids to different ones", () => {
    const { styles } = render((ws) => {
      ws.addConditionalFormatting({
        ref: "A1:A5",
        rules: [
          { type: "cellIs", operator: "greaterThan", formulae: [1], style: HIGHLIGHT, priority: 1 },
          { type: "cellIs", operator: "lessThan", formulae: [0], style: HIGHLIGHT, priority: 2 },
          {
            type: "cellIs",
            operator: "equal",
            formulae: [9],
            style: { font: { bold: true } },
            priority: 3,
          },
        ],
      });
    });
    expect(styles.toXml()).toContain('<dxfs count="2">');
  });

  it("assigns a workbook-monotonic priority when a rule omits one", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("CF");
    ws.addConditionalFormatting({
      ref: "A1:A5",
      rules: [{ type: "duplicateValues" }, { type: "uniqueValues" }],
    });
    const buf = renderWorksheet(ws, {
      styles: createStyles(),
      sharedStrings: createSharedStrings(),
      useSharedStrings: true,
      date1904: false,
      nextPriority: (() => {
        let n = 0;
        return () => ++n;
      })(),
    }).xml;
    // Both rules carry no style, so both intern the same empty dxf (id 0); the priorities advance.
    expect(buf).toContain('type="duplicateValues" dxfId="0" priority="1"');
    expect(buf).toContain('type="uniqueValues" dxfId="0" priority="2"');
  });

  it("sits after </mergeCells> and before <dataValidations> in CT_Worksheet order", () => {
    const { xml } = render((ws) => {
      ws.mergeCells("A1:B1");
      ws.addConditionalFormatting({
        ref: "C3:C9",
        rules: [{ type: "duplicateValues", priority: 1 }],
      });
      ws.dataValidations.add("D1", { type: "list", formulae: ['"a,b"'] });
    });
    expect(xml).toMatch(/<\/mergeCells>.*<conditionalFormatting.*<dataValidations/s);
  });

  it("produces no <conditionalFormatting> block when the sheet has none", () => {
    const { xml } = render((ws) => {
      ws.getCell("A1").value = "plain";
    });
    expect(xml).not.toContain("<conditionalFormatting");
  });
});
