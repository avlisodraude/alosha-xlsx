import { describe, expect, it } from "vitest";
import { InvalidPackageError, UnsupportedFeatureError } from "../../../src/errors.js";
import { Workbook } from "../../../src/model/index.js";
import { CONTENT_TYPE, PART_PATH } from "../../../src/xlsx/parts/constants.js";
import { openPackage } from "../../../src/xlsx/read/package.js";
import { writeWorkbookBuffer } from "../../../src/xlsx/write-workbook.js";
import { writeZipArchive, type ZipEntries } from "../../../src/zip-layer.js";

const enc = new TextEncoder();

/** A minimal, well-formed .xlsx archive built by hand (not via the writer) so tests can mutate one
 *  part at a time and assert precisely how openPackage reacts. */
function minimalPackage(overrides?: Partial<Record<string, string | undefined>>): ZipEntries {
  const parts: Record<string, string> = {
    [PART_PATH.contentTypes]: `<?xml version="1.0"?><Types xmlns="ct">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/xl/workbook.xml" ContentType="${CONTENT_TYPE.workbook}"/>
      <Override PartName="/xl/worksheets/sheet1.xml" ContentType="${CONTENT_TYPE.worksheet}"/>
    </Types>`,
    [PART_PATH.rootRels]: `<?xml version="1.0"?><Relationships xmlns="pr">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
    </Relationships>`,
    [PART_PATH.workbook]: `<?xml version="1.0"?><workbook xmlns="sm"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    [PART_PATH.workbookRels]: `<?xml version="1.0"?><Relationships xmlns="pr">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
    </Relationships>`,
    "xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet xmlns="sm"><sheetData/></worksheet>`,
    "xl/styles.xml": `<?xml version="1.0"?><styleSheet xmlns="sm"/>`,
  };

  if (overrides) {
    for (const [path, xml] of Object.entries(overrides)) {
      if (xml === undefined) delete parts[path];
      else parts[path] = xml;
    }
  }

  const entries: ZipEntries = new Map();
  for (const [path, xml] of Object.entries(parts)) entries.set(path, enc.encode(xml));
  return entries;
}

describe("openPackage — happy path (hand-built minimal package)", () => {
  it("resolves the workbook path from _rels/.rels", () => {
    const parts = openPackage(writeZipArchive(minimalPackage()));
    expect(parts.workbookPath()).toBe("xl/workbook.xml");
  });

  it("exposes entry() and text() over the raw archive", () => {
    const parts = openPackage(writeZipArchive(minimalPackage()));
    expect(parts.text(PART_PATH.workbook)).toContain("<workbook");
    expect(parts.entry(PART_PATH.workbook)).toBeInstanceOf(Uint8Array);
    expect(parts.entry("nonexistent.xml")).toBeUndefined();
    expect(parts.text("nonexistent.xml")).toBeUndefined();
  });

  it("resolves workbook.xml.rels targets relative to xl/", () => {
    const parts = openPackage(writeZipArchive(minimalPackage()));
    const rels = parts.workbookRels();
    expect(rels.get("rId1")).toEqual({
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
      target: "xl/worksheets/sheet1.xml",
    });
    expect(rels.get("rId2")).toEqual({
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
      target: "xl/styles.xml",
    });
  });

  it("resolves ../ relative targets against the workbook's directory", () => {
    const parts = openPackage(
      writeZipArchive(
        minimalPackage({
          [PART_PATH.workbookRels]: `<?xml version="1.0"?><Relationships xmlns="pr">
            <Relationship Id="rId1" Type="theme" Target="../media/image1.png"/>
          </Relationships>`,
        }),
      ),
    );
    expect(parts.workbookRels().get("rId1")?.target).toBe("media/image1.png");
  });

  it("resolves a package-root-absolute target", () => {
    const parts = openPackage(
      writeZipArchive(
        minimalPackage({
          [PART_PATH.workbookRels]: `<?xml version="1.0"?><Relationships xmlns="pr">
            <Relationship Id="rId1" Type="theme" Target="/xl/theme/theme1.xml"/>
          </Relationships>`,
        }),
      ),
    );
    expect(parts.workbookRels().get("rId1")?.target).toBe("xl/theme/theme1.xml");
  });

  it("resolves the workbook part even under a different producer's file name", () => {
    const archive = writeZipArchive(
      minimalPackage({
        [PART_PATH.rootRels]: `<?xml version="1.0"?><Relationships xmlns="pr">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/book.xml"/>
        </Relationships>`,
        [PART_PATH.workbook]: undefined,
        [PART_PATH.contentTypes]: `<?xml version="1.0"?><Types xmlns="ct">
          <Override PartName="/xl/book.xml" ContentType="${CONTENT_TYPE.workbook}"/>
        </Types>`,
        "xl/book.xml": `<?xml version="1.0"?><workbook xmlns="sm"/>`,
        "xl/_rels/book.xml.rels": `<?xml version="1.0"?><Relationships xmlns="pr"/>`,
        [PART_PATH.workbookRels]: undefined,
      }),
    );
    const parts = openPackage(archive);
    expect(parts.workbookPath()).toBe("xl/book.xml");
  });

  it("returns an empty rels map when workbook.xml.rels is absent", () => {
    const parts = openPackage(
      writeZipArchive(minimalPackage({ [PART_PATH.workbookRels]: undefined })),
    );
    expect(parts.workbookRels().size).toBe(0);
  });

  it("partRels resolves a worksheet's own .rels sibling, external targets left as-is", () => {
    const parts = openPackage(
      writeZipArchive(
        minimalPackage({
          "xl/worksheets/_rels/sheet1.xml.rels": `<?xml version="1.0"?><Relationships xmlns="pr">
            <Relationship Id="rId1" Type="hyperlink" Target="https://example.com" TargetMode="External"/>
          </Relationships>`,
        }),
      ),
    );
    const rels = parts.partRels("xl/worksheets/sheet1.xml");
    expect(rels.get("rId1")).toEqual({
      type: "hyperlink",
      target: "https://example.com",
      mode: "External",
    });
  });

  it("partRels resolves an internal target relative to the part's own directory", () => {
    const parts = openPackage(
      writeZipArchive(
        minimalPackage({
          "xl/worksheets/_rels/sheet1.xml.rels": `<?xml version="1.0"?><Relationships xmlns="pr">
            <Relationship Id="rId1" Type="image" Target="../media/image1.png"/>
          </Relationships>`,
        }),
      ),
    );
    const rels = parts.partRels("xl/worksheets/sheet1.xml");
    expect(rels.get("rId1")).toEqual({
      type: "image",
      target: "xl/media/image1.png",
      mode: "Internal",
    });
  });

  it("partRels returns an empty map when the part has no .rels sibling", () => {
    const parts = openPackage(writeZipArchive(minimalPackage()));
    expect(parts.partRels("xl/worksheets/sheet1.xml").size).toBe(0);
  });
});

describe("openPackage — errors", () => {
  it("throws InvalidPackageError on a truncated/corrupt zip", () => {
    expect(() => openPackage(new Uint8Array([1, 2, 3, 4, 5]))).toThrow(InvalidPackageError);
  });

  it("throws InvalidPackageError when _rels/.rels is missing", () => {
    expect(() =>
      openPackage(writeZipArchive(minimalPackage({ [PART_PATH.rootRels]: undefined }))),
    ).toThrow(InvalidPackageError);
  });

  it("throws InvalidPackageError when _rels/.rels has no officeDocument relationship", () => {
    const archive = writeZipArchive(
      minimalPackage({
        [PART_PATH.rootRels]: `<?xml version="1.0"?><Relationships xmlns="pr">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
        </Relationships>`,
      }),
    );
    expect(() => openPackage(archive)).toThrow(InvalidPackageError);
  });

  it("throws InvalidPackageError when the resolved workbook part does not exist", () => {
    const archive = writeZipArchive(minimalPackage({ [PART_PATH.workbook]: undefined }));
    expect(() => openPackage(archive)).toThrow(InvalidPackageError);
  });

  it("throws UnsupportedFeatureError when [Content_Types].xml is missing", () => {
    const archive = writeZipArchive(minimalPackage({ [PART_PATH.contentTypes]: undefined }));
    expect(() => openPackage(archive)).toThrow(UnsupportedFeatureError);
  });

  it("throws UnsupportedFeatureError when the workbook part isn't declared as a spreadsheet", () => {
    const archive = writeZipArchive(
      minimalPackage({
        [PART_PATH.contentTypes]: `<?xml version="1.0"?><Types xmlns="ct">
          <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        </Types>`,
      }),
    );
    expect(() => openPackage(archive)).toThrow(UnsupportedFeatureError);
  });
});

describe("openPackage — real package produced by writeWorkbookBuffer", () => {
  function fixture(): Workbook {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Data");
    ws.getCell("A1").value = "Hello";
    return wb;
  }

  it("opens a real archive and finds every wired-up part", () => {
    const bytes = writeWorkbookBuffer(fixture());
    const parts = openPackage(bytes);

    expect(parts.workbookPath()).toBe(PART_PATH.workbook);
    const rels = parts.workbookRels();
    const targets = new Set([...rels.values()].map((r) => r.target));
    expect(targets).toContain("xl/worksheets/sheet1.xml");
    expect(targets).toContain("xl/styles.xml");
    expect(targets).toContain("xl/sharedStrings.xml");
    expect(targets).toContain("xl/theme/theme1.xml");
    expect(parts.text("xl/worksheets/sheet1.xml")).toContain("<worksheet");
  });
});
