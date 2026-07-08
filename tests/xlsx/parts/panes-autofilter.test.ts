import { describe, expect, it } from "vitest";
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

describe("worksheet writer — panes", () => {
  it("emits a bare <sheetView> with no <pane> when there is no view (default)", () => {
    const xml = render((ws) => {
      ws.getCell("A1").value = "plain";
    });
    expect(xml).toContain('<sheetViews><sheetView workbookViewId="0"/></sheetViews>');
    expect(xml).not.toContain("<pane");
  });

  it("emits a frozen <pane> (golden)", () => {
    const xml = render((ws) => {
      ws.views = [{ state: "frozen", xSplit: 1, ySplit: 1, topLeftCell: "B2" }];
    });
    expect(xml).toContain(
      '<sheetViews><sheetView workbookViewId="0">' +
        '<pane xSplit="1" ySplit="1" topLeftCell="B2" state="frozen"/>' +
        "</sheetView></sheetViews>",
    );
  });

  it("emits a split <pane> including activePane (golden)", () => {
    const xml = render((ws) => {
      ws.views = [
        {
          state: "split",
          xSplit: 2000,
          ySplit: 3000,
          topLeftCell: "C4",
          activePane: "bottomRight",
        },
      ];
    });
    expect(xml).toContain(
      '<pane xSplit="2000" ySplit="3000" topLeftCell="C4" activePane="bottomRight" state="split"/>',
    );
  });

  it("does not emit <pane> for a normal view", () => {
    const xml = render((ws) => {
      ws.views = [{ state: "normal" }];
    });
    expect(xml).not.toContain("<pane");
  });

  it("still honors showGridLines on the sheetView when frozen", () => {
    const xml = render((ws) => {
      ws.properties.showGridLines = false;
      ws.views = [{ state: "frozen", ySplit: 1 }];
    });
    expect(xml).toContain('<sheetView workbookViewId="0" showGridLines="0">');
  });

  it("emits zoomScale and zoomScaleNormal on a plain view", () => {
    const xml = render((ws) => {
      ws.views = [{ zoomScale: 150 }];
    });
    expect(xml).toContain('<sheetView workbookViewId="0" zoomScale="150" zoomScaleNormal="150"/>');
  });

  it("emits zoomScale on a frozen view alongside its pane", () => {
    const xml = render((ws) => {
      ws.views = [{ state: "frozen", ySplit: 1, zoomScale: 75 }];
    });
    expect(xml).toContain(
      '<sheetView workbookViewId="0" zoomScale="75" zoomScaleNormal="75">' +
        '<pane ySplit="1" state="frozen"/>' +
        "</sheetView>",
    );
  });

  it("omits zoomScale when it is the default 100", () => {
    const xml = render((ws) => {
      ws.views = [{ zoomScale: 100 }];
    });
    expect(xml).not.toContain("zoomScale");
  });
});

describe("worksheet writer — autoFilter", () => {
  it("emits <autoFilter ref> (golden)", () => {
    const xml = render((ws) => {
      ws.autoFilter = "A1:D1";
    });
    expect(xml).toContain('<autoFilter ref="A1:D1"/>');
  });

  it("produces no <autoFilter> element when unset", () => {
    const xml = render((ws) => {
      ws.getCell("A1").value = "plain";
    });
    expect(xml).not.toContain("<autoFilter");
  });

  it("sits after <sheetData> and before <mergeCells> in CT_Worksheet order", () => {
    const xml = render((ws) => {
      ws.autoFilter = "A1:C1";
      ws.mergeCells("A1:B1");
    });
    expect(xml).toMatch(/<\/sheetData>.*<autoFilter ref="A1:C1"\/>.*<mergeCells/s);
  });
});
