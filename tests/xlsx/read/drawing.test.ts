// Track D session 3 (reader) — docs/IMAGES-DRAWINGS-SPEC.md §3. Golden inversions of renderDrawing's
// two-cell/one-cell shapes, corner-offset/multi-picture fidelity, the never-throw contract for an
// unresolvable picture, and the typed-error contract for a malformed numeric marker/ext value.

import { describe, expect, it } from "vitest";
import { InvalidDrawingError } from "../../../src/errors.js";
import { type DrawingImage, renderDrawing } from "../../../src/xlsx/parts/drawing.js";
import { parseDrawing } from "../../../src/xlsx/read/drawing.js";

describe("parseDrawing — golden inversions", () => {
  it("inverts a two-cell anchor (from 1,1 to 4,6 -> tl 1,1 / br 3,5, undoing the +1)", () => {
    const xml =
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
      "<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>";

    expect(parseDrawing(xml)).toEqual([
      { embedId: "rId1", anchor: { tl: { col: 1, row: 1 }, br: { col: 3, row: 5 } } },
    ]);
  });

  it("inverts a one-cell anchor's EMU extent back to pixels (÷9525)", () => {
    const images: DrawingImage[] = [
      { embedId: "rId1", anchor: { tl: { col: 2, row: 3 }, ext: { width: 100, height: 50 } } },
    ];
    const xml = renderDrawing(images);
    expect(parseDrawing(xml)).toEqual(images);
  });

  it("preserves two-cell corner offsets through a write -> read round-trip", () => {
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
    expect(parseDrawing(xml)).toEqual(images);
  });

  it("recovers multiple pictures in document order with their own embedId", () => {
    const images: DrawingImage[] = [
      { embedId: "rId1", anchor: { tl: { col: 0, row: 0 }, br: { col: 1, row: 1 } } },
      { embedId: "rId2", anchor: { tl: { col: 3, row: 3 }, ext: { width: 10, height: 10 } } },
    ];
    const xml = renderDrawing(images);
    expect(parseDrawing(xml)).toEqual(images);
  });

  it("skips a picture with no <a:blip r:embed> rather than throwing", () => {
    const xml =
      '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"' +
      ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<xdr:twoCellAnchor editAs="oneCell">' +
      "<xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff>" +
      "<xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>" +
      "<xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff>" +
      "<xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>" +
      '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="Picture 1"/><xdr:cNvPicPr/></xdr:nvPicPr>' +
      "<xdr:blipFill><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>" +
      '<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>' +
      "<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>";

    expect(parseDrawing(xml)).toEqual([]);
  });

  it("skips an <xdr:absoluteAnchor> (deferred per spec §0) without throwing", () => {
    const xml =
      '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">' +
      '<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="100" cy="100"/>' +
      "<xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>";

    expect(parseDrawing(xml)).toEqual([]);
  });

  it("returns an empty array for a rootless/malformed document", () => {
    expect(parseDrawing("<not-a-drawing/>")).toEqual([]);
  });

  it("throws InvalidDrawingError for a non-numeric <xdr:col>, never a silent NaN", () => {
    const xml =
      '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"' +
      ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<xdr:twoCellAnchor editAs="oneCell">' +
      "<xdr:from><xdr:col>x</xdr:col><xdr:colOff>0</xdr:colOff>" +
      "<xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>" +
      "<xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff>" +
      "<xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>" +
      '<xdr:pic><xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill></xdr:pic>' +
      "<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>";

    expect(() => parseDrawing(xml)).toThrow(InvalidDrawingError);
  });

  it("throws InvalidDrawingError for a non-numeric <xdr:ext cx>", () => {
    const xml =
      '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"' +
      ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      "<xdr:oneCellAnchor>" +
      "<xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff>" +
      "<xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>" +
      '<xdr:ext cx="abc" cy="100"/>' +
      '<xdr:pic><xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill></xdr:pic>' +
      "<xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>";

    expect(() => parseDrawing(xml)).toThrow(InvalidDrawingError);
  });
});
