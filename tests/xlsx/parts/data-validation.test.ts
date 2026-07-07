import { describe, expect, it } from "vitest";
import type { DataValidation } from "../../../src/model/index.js";
import { Workbook } from "../../../src/model/index.js";
import { renderWorksheet } from "../../../src/xlsx/parts/worksheet.js";
import { createSharedStrings } from "../../../src/xlsx/shared-strings.js";
import { createStyles } from "../../../src/xlsx/styles.js";

/** Render a single-sheet worksheet through the real pools, returning its xml. */
function render(build: (ws: ReturnType<Workbook["addWorksheet"]>) => void): string {
  const ws = new Workbook().addWorksheet("Data");
  build(ws);
  return renderWorksheet(ws, {
    styles: createStyles(),
    sharedStrings: createSharedStrings(),
    useSharedStrings: true,
    date1904: false,
  }).xml;
}

describe("worksheet writer — data validations", () => {
  it("emits an inline list dropdown (golden)", () => {
    const xml = render((ws) => {
      ws.getCell("A1").dataValidation = {
        type: "list",
        allowBlank: true,
        showInputMessage: true,
        showErrorMessage: true,
        formulae: ['"Yes,No,Maybe"'],
      };
    });
    expect(xml).toContain(
      '<dataValidations count="1">' +
        '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1">' +
        '<formula1>"Yes,No,Maybe"</formula1>' +
        "</dataValidation>" +
        "</dataValidations>",
    );
  });

  it("emits a same-sheet range list dropdown (golden)", () => {
    const xml = render((ws) => {
      ws.dataValidations.add("A1:A10", { type: "list", formulae: ["$E$1:$E$5"] });
    });
    expect(xml).toContain(
      '<dataValidation type="list" sqref="A1:A10">' +
        "<formula1>$E$1:$E$5</formula1>" +
        "</dataValidation>",
    );
  });

  it("emits a decimal between with two formulae (golden)", () => {
    const xml = render((ws) => {
      ws.dataValidations.add("B2", {
        type: "decimal",
        operator: "between",
        formulae: [0.5, 9.5],
      });
    });
    expect(xml).toContain(
      '<dataValidation type="decimal" operator="between" sqref="B2">' +
        "<formula1>0.5</formula1>" +
        "<formula2>9.5</formula2>" +
        "</dataValidation>",
    );
  });

  it("omits formula2 for a one-sided operator", () => {
    const xml = render((ws) => {
      ws.dataValidations.add("C3", {
        type: "whole",
        operator: "greaterThan",
        formulae: [10, 999],
      });
    });
    expect(xml).toContain(
      '<dataValidation type="whole" operator="greaterThan" sqref="C3">' +
        "<formula1>10</formula1>" +
        "</dataValidation>",
    );
    expect(xml).not.toContain("<formula2>");
  });

  it("emits a custom rule with no operator (golden)", () => {
    const xml = render((ws) => {
      ws.dataValidations.add("D4", { type: "custom", formulae: ["ISNUMBER(D4)"] });
    });
    expect(xml).toContain(
      '<dataValidation type="custom" sqref="D4"><formula1>ISNUMBER(D4)</formula1></dataValidation>',
    );
    expect(xml).not.toContain("operator=");
  });

  it("omits the operator attribute for list and custom types", () => {
    const xml = render((ws) => {
      ws.dataValidations.add("A1", {
        type: "list",
        operator: "between",
        formulae: ['"a,b"'],
      } as DataValidation);
    });
    expect(xml).not.toContain("operator=");
  });

  it("passes the message/error attributes through (golden)", () => {
    const xml = render((ws) => {
      ws.dataValidations.add("E5", {
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
    expect(xml).toContain(
      '<dataValidation type="textLength" operator="lessThanOrEqual" showErrorMessage="1"' +
        ' errorStyle="warning" errorTitle="Too long" error="Keep it short"' +
        ' promptTitle="Note" prompt="Enter a short code" sqref="E5">' +
        "<formula1>8</formula1>" +
        "</dataValidation>",
    );
  });

  it('inverts showDropDown: model false → attr "1"', () => {
    const xml = render((ws) => {
      ws.dataValidations.add("F6", {
        type: "list",
        showDropDown: false,
        formulae: ['"a,b"'],
      });
    });
    expect(xml).toContain('<dataValidation type="list" showDropDown="1" sqref="F6">');
  });

  it("omits showDropDown when the model shows the dropdown (true or unset)", () => {
    const shown = render((ws) => {
      ws.dataValidations.add("F6", { type: "list", showDropDown: true, formulae: ['"a,b"'] });
    });
    expect(shown).not.toContain("showDropDown");

    const unset = render((ws) => {
      ws.dataValidations.add("F7", { type: "list", formulae: ['"a,b"'] });
    });
    expect(unset).not.toContain("showDropDown");
  });

  it("sits after <mergeCells> and before <hyperlinks> in CT_Worksheet order", () => {
    const xml = render((ws) => {
      ws.mergeCells("A1:B1");
      ws.dataValidations.add("C3", { type: "list", formulae: ['"a,b"'] });
      ws.getCell("A5").value = { text: "site", hyperlink: "https://example.com" };
    });
    expect(xml).toMatch(/<\/mergeCells>.*<dataValidations.*<hyperlinks>/s);
  });

  it("emits count matching the number of validations and no rels effect", () => {
    const xml = render((ws) => {
      ws.dataValidations.add("A1", { type: "list", formulae: ['"a,b"'] });
      ws.dataValidations.add("B1", { type: "custom", formulae: ["TRUE()"] });
    });
    expect(xml).toContain('<dataValidations count="2">');
  });

  it("produces no <dataValidations> block when the sheet has none", () => {
    const xml = render((ws) => {
      ws.getCell("A1").value = "plain";
    });
    expect(xml).not.toContain("<dataValidations");
  });
});
