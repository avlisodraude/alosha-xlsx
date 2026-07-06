import { describe, expect, it } from "vitest";
import { Workbook } from "../../../src/model/index.js";
import { renderAppProps } from "../../../src/xlsx/parts/app-props.js";
import { renderCoreProps, toW3CDTF } from "../../../src/xlsx/parts/core-props.js";
import { applyAppProps, applyCoreProps, fromW3CDTF } from "../../../src/xlsx/read/props.js";

describe("fromW3CDTF", () => {
  it("inverts toW3CDTF", () => {
    const date = new Date("2026-01-02T03:04:05.000Z");
    expect(fromW3CDTF(toW3CDTF(date))).toEqual(date);
  });
});

describe("applyCoreProps", () => {
  it("is a no-op when the part is empty/absent", () => {
    const wb = new Workbook();
    const before = { ...wb };
    applyCoreProps(wb, "");
    expect(wb.creator).toBe(before.creator);
    expect(wb.title).toBeUndefined();
  });

  it("applies creator, created and does not set unset fields", () => {
    const wb = new Workbook();
    const xml =
      '<cp:coreProperties xmlns:cp="cp" xmlns:dc="dc" xmlns:dcterms="dcterms" xmlns:xsi="xsi">' +
      "<dc:creator>Eduardo</dc:creator>" +
      '<dcterms:created xsi:type="dcterms:W3CDTF">2026-01-02T03:04:05Z</dcterms:created>' +
      '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-02-03T04:05:06Z</dcterms:modified>' +
      "</cp:coreProperties>";
    applyCoreProps(wb, xml);
    expect(wb.creator).toBe("Eduardo");
    expect(wb.created).toEqual(new Date("2026-01-02T03:04:05Z"));
    expect(wb.modified).toEqual(new Date("2026-02-03T04:05:06Z"));
    expect(wb.title).toBeUndefined();
  });

  it("applies every optional metadata field", () => {
    const wb = new Workbook();
    const xml =
      '<cp:coreProperties xmlns:cp="cp" xmlns:dc="dc" xmlns:dcterms="dcterms">' +
      "<dc:title>My Title</dc:title>" +
      "<dc:subject>My Subject</dc:subject>" +
      "<dc:description>My Description</dc:description>" +
      "<cp:keywords>a, b</cp:keywords>" +
      "<cp:category>Reports</cp:category>" +
      "<cp:lastModifiedBy>Ana</cp:lastModifiedBy>" +
      "</cp:coreProperties>";
    applyCoreProps(wb, xml);
    expect(wb.title).toBe("My Title");
    expect(wb.subject).toBe("My Subject");
    expect(wb.description).toBe("My Description");
    expect(wb.keywords).toBe("a, b");
    expect(wb.category).toBe("Reports");
    expect(wb.lastModifiedBy).toBe("Ana");
  });

  it("round-trips renderCoreProps' own output", () => {
    const wb = new Workbook();
    wb.creator = "Eduardo";
    wb.title = "T";
    wb.created = new Date("2026-01-02T03:04:05.000Z");
    wb.modified = new Date("2026-02-03T04:05:06.000Z");
    const xml = renderCoreProps(wb);

    const parsed = new Workbook();
    applyCoreProps(parsed, xml);
    expect(parsed.creator).toBe("Eduardo");
    expect(parsed.title).toBe("T");
    expect(parsed.created).toEqual(wb.created);
    expect(parsed.modified).toEqual(wb.modified);
  });
});

describe("applyAppProps", () => {
  it("is a no-op when the part is empty/absent", () => {
    const wb = new Workbook();
    applyAppProps(wb, "");
    expect(wb.company).toBeUndefined();
    expect(wb.manager).toBeUndefined();
  });

  it("applies Company/Manager, skipping over the HeadingPairs/TitlesOfParts vectors", () => {
    const wb = new Workbook();
    const xml = renderAppProps(
      Object.assign(new Workbook(), { company: "Acme", manager: "Boss" }),
      ["Data", "Sheet2"],
    );
    applyAppProps(wb, xml);
    expect(wb.company).toBe("Acme");
    expect(wb.manager).toBe("Boss");
  });

  it("leaves Company/Manager unset when absent from the part", () => {
    const wb = new Workbook();
    const xml = renderAppProps(new Workbook(), ["Data"]);
    applyAppProps(wb, xml);
    expect(wb.company).toBeUndefined();
    expect(wb.manager).toBeUndefined();
  });
});
