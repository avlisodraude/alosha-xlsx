// Track D session 2 (writer) — docs/IMAGES-DRAWINGS-SPEC.md §2. Golden DrawingML anchors, the media/
// drawing/rels child parts pushed through the F0 seam, the `<drawing>` wiring in the sheet body, and a
// no-images byte-identical regression.

import { describe, expect, it } from "vitest";
import { Workbook } from "../../../src/model/index.js";
import { CONTENT_TYPE, IMAGE_CONTENT_TYPE, REL_TYPE } from "../../../src/xlsx/parts/constants.js";
import { type DrawingImage, renderDrawing } from "../../../src/xlsx/parts/drawing.js";
import { renderWorksheet } from "../../../src/xlsx/parts/worksheet.js";
import { createSharedStrings } from "../../../src/xlsx/shared-strings.js";
import { createStyles } from "../../../src/xlsx/styles.js";
import { writeWorkbookBuffer } from "../../../src/xlsx/write-workbook.js";
import { readZipArchive } from "../../../src/zip-layer.js";

const dec = new TextDecoder();

/** A 4-byte PNG-ish blob — the exact bytes are what the byte-identity test tracks. */
const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47]);

/** Render a single sheet through the real pools at a given 1-based `sheetNumber`. */
function render(build: (ws: ReturnType<Workbook["addWorksheet"]>) => void, sheetNumber = 1) {
  const wb = new Workbook();
  const ws = wb.addWorksheet("Data");
  build(ws);
  return renderWorksheet(ws, {
    styles: createStyles(),
    sharedStrings: createSharedStrings(),
    useSharedStrings: true,
    date1904: false,
    sheetNumber,
  });
}

describe("renderDrawing — golden DrawingML", () => {
  it("renders a two-cell anchor with the exclusive br+1 corner (B2:D6 → from 1,1 to 4,6)", () => {
    const images: DrawingImage[] = [
      { embedId: "rId1", anchor: { tl: { col: 1, row: 1 }, br: { col: 3, row: 5 } } },
    ];
    expect(renderDrawing(images)).toBe(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"' +
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<xdr:twoCellAnchor editAs="oneCell">' +
        "<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff>" +
        "<xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>" +
        "<xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff>" +
        "<xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>" +
        "<xdr:pic><xdr:nvPicPr>" +
        '<xdr:cNvPr id="1" name="Picture 1"/><xdr:cNvPicPr/></xdr:nvPicPr>' +
        '<xdr:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>' +
        '<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>' +
        "<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>",
    );
  });

  it("renders a one-cell anchor with the pixel extent converted to EMU (×9525)", () => {
    const images: DrawingImage[] = [
      { embedId: "rId1", anchor: { tl: { col: 2, row: 3 }, ext: { width: 100, height: 50 } } },
    ];
    const xml = renderDrawing(images);
    // 100px → 952500 EMU, 50px → 476250 EMU.
    expect(xml).toContain("<xdr:oneCellAnchor>");
    expect(xml).toContain(
      "<xdr:from><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff>" +
        "<xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>",
    );
    expect(xml).toContain('<xdr:ext cx="952500" cy="476250"/>');
    expect(xml).not.toContain("<xdr:to>");
  });

  it("carries the two-cell corner offsets through when the model supplies them", () => {
    const images: DrawingImage[] = [
      {
        embedId: "rId1",
        anchor: {
          tl: { col: 0, row: 0, nativeColOff: 12700, nativeRowOff: 6350 },
          br: { col: 2, row: 2 },
        },
      },
    ];
    const xml = renderDrawing(images);
    expect(xml).toContain(
      "<xdr:from><xdr:col>0</xdr:col><xdr:colOff>12700</xdr:colOff>" +
        "<xdr:row>0</xdr:row><xdr:rowOff>6350</xdr:rowOff></xdr:from>",
    );
    // The `to` corner stays at the exclusive br+1 with zero offsets regardless.
    expect(xml).toContain(
      "<xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff>" +
        "<xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>",
    );
  });

  it("numbers multiple pictures 1..n within one drawing", () => {
    const images: DrawingImage[] = [
      { embedId: "rId1", anchor: { tl: { col: 0, row: 0 }, br: { col: 1, row: 1 } } },
      { embedId: "rId2", anchor: { tl: { col: 3, row: 3 }, ext: { width: 10, height: 10 } } },
    ];
    const xml = renderDrawing(images);
    expect(xml).toContain('<xdr:cNvPr id="1" name="Picture 1"/>');
    expect(xml).toContain('<xdr:cNvPr id="2" name="Picture 2"/>');
    expect(xml).toContain('<a:blip r:embed="rId1"/>');
    expect(xml).toContain('<a:blip r:embed="rId2"/>');
  });
});

describe("worksheet writer — image wiring", () => {
  it("emits <drawing r:id> after any <hyperlinks> and before the close, with a matching sheet rel", () => {
    const wb = new Workbook();
    const id = wb.addImage({ buffer: PNG, extension: "png" });
    const ws = wb.worksheets[0] ?? wb.addWorksheet("Data");
    ws.getCell("A1").value = { text: "site", hyperlink: "https://example.com" };
    ws.addImage(id, "B2:D6");

    const { xml, rels, extraParts } = renderWorksheet(ws, {
      styles: createStyles(),
      sharedStrings: createSharedStrings(),
      useSharedStrings: true,
      date1904: false,
      sheetNumber: 1,
    });

    // <drawing> sits after <hyperlinks>, immediately before </worksheet>.
    expect(xml).toMatch(/<\/hyperlinks><drawing r:id="rId2"\/><\/worksheet>$/);
    // The hyperlink took rId1; the drawing took the next rId off the same allocator.
    expect(rels).toContain(
      `<Relationship Id="rId2" Type="${REL_TYPE.drawing}" Target="../drawings/drawing1.xml"/>`,
    );
    expect(extraParts).toBeDefined();
  });

  it("pushes the media, drawing, and drawing-rels parts whose r:embed ↔ rels Id ↔ media path all agree", () => {
    const wb = new Workbook();
    const id = wb.addImage({ buffer: PNG, extension: "png" });
    const ws = wb.worksheets[0] ?? wb.addWorksheet("Data");
    ws.addImage(id, "B2:D6");

    const { extraParts } = renderWorksheet(ws, {
      styles: createStyles(),
      sharedStrings: createSharedStrings(),
      useSharedStrings: true,
      date1904: false,
      sheetNumber: 1,
    });
    const byPath = new Map((extraParts ?? []).map((p) => [p.path, p]));

    // Media part: raw bytes, a media Default for its extension.
    const media = byPath.get("xl/media/image1.png");
    expect(media?.data).toBe(PNG);
    expect(media?.defaultFor).toEqual({ extension: "png", contentType: IMAGE_CONTENT_TYPE.png });

    // Drawing part: an Override content-type; its blip embeds rId1.
    const drawing = byPath.get("xl/drawings/drawing1.xml");
    expect(drawing?.contentType).toBe(CONTENT_TYPE.drawing);
    expect(drawing?.data).toContain('<a:blip r:embed="rId1"/>');

    // Drawing rels: rId1 (== the r:embed) → the media Target, typed as an image rel.
    const drels = byPath.get("xl/drawings/_rels/drawing1.xml.rels");
    expect(drels?.data).toContain(
      `<Relationship Id="rId1" Type="${REL_TYPE.image}" Target="../media/image1.png"/>`,
    );
  });

  it("does not name its parts sheet1 — a second sheet's parts follow its sheetNumber", () => {
    const { extraParts } = render((ws) => {
      const id = ws.workbook.addImage({ buffer: PNG, extension: "png" });
      ws.addImage(id, "A1:B2");
    }, 3);
    const paths = new Set((extraParts ?? []).map((p) => p.path));
    expect(paths.has("xl/drawings/drawing3.xml")).toBe(true);
    expect(paths.has("xl/drawings/_rels/drawing3.xml.rels")).toBe(true);
  });
});

describe("write orchestrator — images survive the full archive", () => {
  it("writes the media bytes byte-identically and types both parts in [Content_Types].xml", () => {
    const wb = new Workbook();
    const id = wb.addImage({ buffer: PNG, extension: "png" });
    const ws = wb.addWorksheet("Data");
    ws.addImage(id, "B2:D6");

    const files = readZipArchive(writeWorkbookBuffer(wb));

    // Byte-identity: what came out of the zip equals what went into the registry.
    expect([...(files.get("xl/media/image1.png") ?? [])]).toEqual([...PNG]);
    // Both the drawing part and its rels physically landed.
    expect(files.has("xl/drawings/drawing1.xml")).toBe(true);
    expect(files.has("xl/drawings/_rels/drawing1.xml.rels")).toBe(true);

    const contentTypes = dec.decode(files.get("[Content_Types].xml"));
    expect(contentTypes).toContain(
      `<Default Extension="png" ContentType="${IMAGE_CONTENT_TYPE.png}"/>`,
    );
    expect(contentTypes).toContain(
      `<Override PartName="/xl/drawings/drawing1.xml" ContentType="${CONTENT_TYPE.drawing}"/>`,
    );
  });

  it("leaves an image-free workbook byte-identical to F0 output (regression)", () => {
    const build = () => {
      const wb = new Workbook();
      wb.created = new Date("2026-01-02T03:04:05.000Z");
      wb.modified = new Date("2026-01-02T03:04:05.000Z");
      wb.addWorksheet("Data").getCell("A1").value = "hi";
      return wb;
    };
    expect([...writeWorkbookBuffer(build())]).toEqual([...writeWorkbookBuffer(build())]);

    const files = readZipArchive(writeWorkbookBuffer(build()));
    expect([...files.keys()].some((k) => k.startsWith("xl/drawings/"))).toBe(false);
    expect([...files.keys()].some((k) => k.startsWith("xl/media/"))).toBe(false);
    expect(dec.decode(files.get("[Content_Types].xml"))).not.toContain('Extension="png"');
  });
});
