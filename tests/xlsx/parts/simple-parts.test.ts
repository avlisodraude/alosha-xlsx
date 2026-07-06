import { describe, expect, it } from "vitest";
import { Workbook } from "../../../src/model/index.js";
import { renderAppProps } from "../../../src/xlsx/parts/app-props.js";
import type { SheetRel, WorkbookRel } from "../../../src/xlsx/parts/constants.js";
import { REL_TYPE } from "../../../src/xlsx/parts/constants.js";
import { renderContentTypes } from "../../../src/xlsx/parts/content-types.js";
import { renderCoreProps, toW3CDTF } from "../../../src/xlsx/parts/core-props.js";
import { renderRootRels } from "../../../src/xlsx/parts/root-rels.js";
import { renderTheme } from "../../../src/xlsx/parts/theme.js";
import { renderWorkbook } from "../../../src/xlsx/parts/workbook.js";
import { renderWorkbookRels } from "../../../src/xlsx/parts/workbook-rels.js";

const sheetRels: SheetRel[] = [
  { rId: "rId1", fileName: "sheet1.xml", sheetId: 1, name: "Data", state: "visible" },
  { rId: "rId2", fileName: "sheet2.xml", sheetId: 2, name: "Hidden < & >", state: "hidden" },
];

describe("content-types", () => {
  it("declares defaults, both worksheets, and sharedStrings only when present", () => {
    const withSst = renderContentTypes({
      worksheetFiles: ["sheet1.xml", "sheet2.xml"],
      hasSharedStrings: true,
    });
    expect(withSst).toContain('<Default Extension="rels"');
    expect(withSst).toContain('<Default Extension="xml" ContentType="application/xml"/>');
    expect(withSst).toContain('PartName="/xl/workbook.xml"');
    expect(withSst).toContain('PartName="/xl/worksheets/sheet1.xml"');
    expect(withSst).toContain('PartName="/xl/worksheets/sheet2.xml"');
    expect(withSst).toContain('PartName="/xl/sharedStrings.xml"');
    expect(withSst).toContain('PartName="/docProps/core.xml"');

    const noSst = renderContentTypes({ worksheetFiles: ["sheet1.xml"], hasSharedStrings: false });
    expect(noSst).not.toContain("sharedStrings.xml");
  });
});

describe("root rels", () => {
  it("links workbook, core and app with local rIds", () => {
    const xml = renderRootRels();
    expect(xml).toContain(`Id="rId1" Type="${REL_TYPE.officeDocument}" Target="xl/workbook.xml"`);
    expect(xml).toContain(`Id="rId2" Type="${REL_TYPE.coreProps}" Target="docProps/core.xml"`);
    expect(xml).toContain(`Id="rId3" Type="${REL_TYPE.extProps}" Target="docProps/app.xml"`);
  });
});

describe("workbook + rels rId agreement", () => {
  it("emits sheets whose r:id, state and name round-trip", () => {
    const wb = new Workbook();
    wb.properties.date1904 = true;
    const xml = renderWorkbook(wb, sheetRels);
    expect(xml).toContain('<workbookPr date1904="1"/>');
    expect(xml).toContain('<sheet name="Data" sheetId="1" r:id="rId1"/>');
    // Hidden sheet keeps its state; special chars in the name are attribute-escaped.
    expect(xml).toContain('name="Hidden &lt; &amp; &gt;"');
    expect(xml).toContain('state="hidden"');
    expect(xml).toContain('r:id="rId2"');
  });

  it("workbook.xml.rels uses the same sheet rIds plus styles/theme", () => {
    const rels: WorkbookRel[] = [
      { rId: "rId1", type: REL_TYPE.worksheet, target: "worksheets/sheet1.xml" },
      { rId: "rId2", type: REL_TYPE.worksheet, target: "worksheets/sheet2.xml" },
      { rId: "rId3", type: REL_TYPE.styles, target: "styles.xml" },
      { rId: "rId4", type: REL_TYPE.theme, target: "theme/theme1.xml" },
    ];
    const xml = renderWorkbookRels(rels);
    expect(xml).toContain(`Id="rId1" Type="${REL_TYPE.worksheet}" Target="worksheets/sheet1.xml"`);
    expect(xml).toContain(`Id="rId3" Type="${REL_TYPE.styles}" Target="styles.xml"`);
    expect(xml).toContain(`Id="rId4" Type="${REL_TYPE.theme}" Target="theme/theme1.xml"`);
  });
});

describe("core + app props", () => {
  it("core.xml emits W3CDTF dates (no milliseconds) and set metadata only", () => {
    const wb = new Workbook();
    wb.created = new Date("2026-01-02T03:04:05.678Z");
    wb.modified = new Date("2026-02-03T04:05:06.000Z");
    wb.creator = "Eduardo";
    const xml = renderCoreProps(wb);
    expect(toW3CDTF(wb.created)).toBe("2026-01-02T03:04:05Z");
    expect(xml).toContain("<dc:creator>Eduardo</dc:creator>");
    expect(xml).toContain(
      '<dcterms:created xsi:type="dcterms:W3CDTF">2026-01-02T03:04:05Z</dcterms:created>',
    );
    expect(xml).not.toContain("<dc:title>");
  });

  it("app.xml lists the worksheet titles in a vt:vector", () => {
    const wb = new Workbook();
    const xml = renderAppProps(wb, ["Data", "Sheet2"]);
    expect(xml).toContain("<Application>Microsoft Excel</Application>");
    expect(xml).toContain('<vt:vector size="2" baseType="lpstr">');
    expect(xml).toContain("<vt:lpstr>Data</vt:lpstr>");
    expect(xml).toContain("<vt:lpstr>Sheet2</vt:lpstr>");
  });
});

describe("theme", () => {
  it("ships a well-formed default Office theme with the clrScheme", () => {
    const xml = renderTheme();
    expect(xml).toContain(
      '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
    );
    expect(xml).toContain('<a:clrScheme name="Office">');
    expect(xml).toContain("</a:theme>");
  });
});
