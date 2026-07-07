// F0 (multi-part worksheet plumbing, docs/MULTIPART-PLUMBING-SPEC.md §7): synthetic coverage of the
// generalised seam — mixed worksheet rels, pluggable content-types, and the collect/assemble path —
// with no real feature yet. Proves Tracks C (comments) and D (images) can push child parts through it.

import { describe, expect, it } from "vitest";
import { Workbook } from "../../../src/model/index.js";
import {
  CONTENT_TYPE,
  IMAGE_CONTENT_TYPE,
  REL_TYPE,
  type WorksheetRel,
} from "../../../src/xlsx/parts/constants.js";
import { renderContentTypes } from "../../../src/xlsx/parts/content-types.js";
import type { RenderedWorksheet } from "../../../src/xlsx/parts/worksheet.js";
import { renderWorksheetRels } from "../../../src/xlsx/parts/worksheet-rels.js";
import { collectWorksheetParts, writeWorkbookBuffer } from "../../../src/xlsx/write-workbook.js";
import { readZipArchive, writeZipArchive, type ZipEntries } from "../../../src/zip-layer.js";

const dec = new TextDecoder();

describe("renderWorksheetRels — mixed rel types", () => {
  it("renders a hyperlink (with TargetMode) and an internal drawing (without) in one .rels", () => {
    const rels: WorksheetRel[] = [
      { rId: "rId1", type: REL_TYPE.hyperlink, target: "https://example.com", mode: "External" },
      { rId: "rId2", type: REL_TYPE.drawing, target: "../drawings/drawing1.xml" },
    ];
    const xml = renderWorksheetRels(rels);

    expect(xml).toContain(
      `<Relationship Id="rId1" Type="${REL_TYPE.hyperlink}" Target="https://example.com" TargetMode="External"/>`,
    );
    // The internal drawing rel carries its own Type and NO TargetMode.
    expect(xml).toContain(
      `<Relationship Id="rId2" Type="${REL_TYPE.drawing}" Target="../drawings/drawing1.xml"/>`,
    );
    // TargetMode appears exactly once — only on the external hyperlink.
    expect(xml.match(/TargetMode=/g)).toHaveLength(1);
  });
});

describe("renderContentTypes — pluggable defaults + overrides", () => {
  it("collapses duplicate extra defaults, adds an override, and keeps all defaults before overrides", () => {
    const xml = renderContentTypes({
      worksheetFiles: ["sheet1.xml"],
      hasSharedStrings: false,
      extraDefaults: [
        { extension: "png", contentType: IMAGE_CONTENT_TYPE.png },
        { extension: "png", contentType: IMAGE_CONTENT_TYPE.png },
      ],
      extraOverrides: [
        { partName: "/xl/drawings/drawing1.xml", contentType: CONTENT_TYPE.drawing },
      ],
    });

    // Duplicate png yields exactly one <Default Extension="png">.
    expect(xml.match(/<Default Extension="png"/g)).toHaveLength(1);
    expect(xml).toContain(`<Default Extension="png" ContentType="${IMAGE_CONTENT_TYPE.png}"/>`);
    // The extra override lands.
    expect(xml).toContain(
      `<Override PartName="/xl/drawings/drawing1.xml" ContentType="${CONTENT_TYPE.drawing}"/>`,
    );
    // Schema invariant: every <Default> precedes every <Override>.
    expect(xml.lastIndexOf("<Default ")).toBeLessThan(xml.indexOf("<Override "));
  });

  it("never re-declares the base rels/xml defaults from an extra", () => {
    const xml = renderContentTypes({
      worksheetFiles: ["sheet1.xml"],
      hasSharedStrings: false,
      extraDefaults: [{ extension: "XML", contentType: "application/xml" }],
    });
    expect(xml.match(/<Default Extension="xml"/gi)).toHaveLength(1);
  });
});

describe("collectWorksheetParts", () => {
  const enc = new TextEncoder();

  it("routes an XML part and a binary media part to media, overrides, and defaults", () => {
    const rendered: RenderedWorksheet[] = [
      {
        xml: "<worksheet/>",
        extraParts: [
          { path: "xl/drawings/drawing1.xml", data: "<xml/>", contentType: CONTENT_TYPE.drawing },
          {
            path: "xl/media/image1.png",
            data: new Uint8Array([1, 2, 3]),
            defaultFor: { extension: "png", contentType: IMAGE_CONTENT_TYPE.png },
          },
        ],
      },
    ];

    const collected = collectWorksheetParts(rendered, enc);

    // String data is encoded; raw bytes pass through untouched.
    expect(dec.decode(collected.media.get("xl/drawings/drawing1.xml"))).toBe("<xml/>");
    expect([...(collected.media.get("xl/media/image1.png") ?? [])]).toEqual([1, 2, 3]);
    // The XML part contributes an <Override>; the media part a <Default>.
    expect(collected.extraOverrides).toEqual([
      { partName: "/xl/drawings/drawing1.xml", contentType: CONTENT_TYPE.drawing },
    ]);
    expect(collected.extraDefaults).toEqual([
      { extension: "png", contentType: IMAGE_CONTENT_TYPE.png },
    ]);
  });

  it("ignores sheets with no extra parts", () => {
    const collected = collectWorksheetParts([{ xml: "<worksheet/>" }], enc);
    expect(collected.media.size).toBe(0);
    expect(collected.extraDefaults).toHaveLength(0);
    expect(collected.extraOverrides).toHaveLength(0);
  });
});

describe("write orchestrator — extra parts land in the real archive", () => {
  const enc = new TextEncoder();

  it("assembles synthetic child parts into an actual zip with a matching [Content_Types].xml", () => {
    // No public feature emits extraParts in F0, so drive the exact assembly the orchestrator (§6)
    // would: collect a hand-built RenderedWorksheet, fold its defaults/overrides into the
    // content-types, write the media into the entries, zip, and read it back.
    const rendered: RenderedWorksheet[] = [
      {
        xml: "<worksheet/>",
        extraParts: [
          { path: "xl/drawings/drawing1.xml", data: "<xml/>", contentType: CONTENT_TYPE.drawing },
          {
            path: "xl/media/image1.png",
            data: new Uint8Array([1, 2, 3]),
            defaultFor: { extension: "png", contentType: IMAGE_CONTENT_TYPE.png },
          },
        ],
      },
    ];
    const collected = collectWorksheetParts(rendered, enc);

    const entries: ZipEntries = new Map();
    entries.set(
      "[Content_Types].xml",
      enc.encode(
        renderContentTypes({
          worksheetFiles: ["sheet1.xml"],
          hasSharedStrings: false,
          extraDefaults: collected.extraDefaults,
          extraOverrides: collected.extraOverrides,
        }),
      ),
    );
    for (const [path, bytes] of collected.media) entries.set(path, bytes);

    const files = readZipArchive(writeZipArchive(entries));

    // Both child parts are physically present, at their declared paths and with their exact bytes.
    expect(dec.decode(files.get("xl/drawings/drawing1.xml"))).toBe("<xml/>");
    expect([...(files.get("xl/media/image1.png") ?? [])]).toEqual([1, 2, 3]);
    // …and [Content_Types].xml types both: a <Default> for png, an <Override> for the drawing.
    const contentTypes = dec.decode(files.get("[Content_Types].xml"));
    expect(contentTypes).toContain(
      `<Default Extension="png" ContentType="${IMAGE_CONTENT_TYPE.png}"/>`,
    );
    expect(contentTypes).toContain(
      `<Override PartName="/xl/drawings/drawing1.xml" ContentType="${CONTENT_TYPE.drawing}"/>`,
    );
  });

  it("leaves a normal workbook's archive free of any phantom child parts (regression)", () => {
    const wb = new Workbook();
    wb.addWorksheet("Data").getCell("A1").value = "hi";

    const files = readZipArchive(writeWorkbookBuffer(wb));
    const contentTypes = dec.decode(files.get("[Content_Types].xml"));

    expect(files.has("xl/media/image1.png")).toBe(false);
    expect(contentTypes).not.toContain('Extension="png"');
    expect(contentTypes).not.toContain("/xl/drawings/");
  });
});
