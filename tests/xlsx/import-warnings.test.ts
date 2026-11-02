import { describe, expect, it } from "vitest";
import {
  type ImportWarning,
  ImportWarningError,
  readWorkbookBuffer,
  Workbook,
  writeWorkbookBuffer,
} from "../../src/index.js";
import { readZipArchive, writeZipArchive } from "../../src/zip-layer.js";

function fixture(change?: (parts: Map<string, Uint8Array>) => void): Uint8Array {
  const wb = new Workbook();
  wb.addWorksheet("S").getCell("A1").value = 42;
  const parts = readZipArchive(writeWorkbookBuffer(wb));
  change?.(parts);
  return writeZipArchive(parts);
}
function edit(
  parts: Map<string, Uint8Array>,
  path: string,
  replace: (xml: string) => string,
): void {
  parts.set(path, new TextEncoder().encode(replace(new TextDecoder().decode(parts.get(path)))));
}

describe("import warnings and strict mode", () => {
  it("locates unsupported parts by content type even at unconventional paths", () => {
    const bytes = fixture((parts) => {
      parts.set("custom/report.xml", new TextEncoder().encode("<chart/>"));
      edit(parts, "[Content_Types].xml", (xml) =>
        xml.replace(
          "</Types>",
          '<Override PartName="/custom/report.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>',
        ),
      );
    });
    const warnings: ImportWarning[] = [];
    readWorkbookBuffer(bytes, { onWarning: (warning) => warnings.push(warning) });
    expect(warnings).toContainEqual(
      expect.objectContaining({ part: "custom/report.xml", feature: "charts" }),
    );
  });
  it("keeps default imports tolerant and reports known discarded content only when requested", () => {
    const bytes = fixture((parts) => {
      parts.set("xl/charts/chart1.xml", new TextEncoder().encode("<chart/>"));
      parts.set("xl/tables/table1.xml", new TextEncoder().encode("<table/>"));
      edit(parts, "xl/workbook.xml", (xml) =>
        xml.replace(
          "</workbook>",
          '<definedNames><definedName name="Amount">S!A1</definedName></definedNames></workbook>',
        ),
      );
      edit(parts, "xl/worksheets/sheet1.xml", (xml) =>
        xml.replace(
          "</worksheet>",
          '<conditionalFormatting sqref="A1"><cfRule type="timePeriod" priority="1"/></conditionalFormatting></worksheet>',
        ),
      );
    });
    expect(readWorkbookBuffer(bytes).worksheets[0]?.getCell("A1").value).toBe(42);
    const warnings: ImportWarning[] = [];
    expect(
      readWorkbookBuffer(bytes, {
        onWarning: (warning) => warnings.push(warning),
      }).worksheets[0]?.getCell("A1").value,
    ).toBe(42);
    expect(warnings.map((warning) => warning.feature)).toEqual(
      expect.arrayContaining(["charts", "tables", "definedNames", "conditional formatting rule"]),
    );
    expect(
      warnings.every(
        (warning) =>
          Object.isFrozen(warning) && warning.part.length > 0 && warning.message.length > 0,
      ),
    ).toBe(true);
  });

  it("calls the observer before throwing a typed strict warning and preserves the receiver", async () => {
    const bytes = fixture((parts) => {
      parts.set("xl/charts/chart1.xml", new TextEncoder().encode("<chart/>"));
    });
    const warnings: ImportWarning[] = [];
    const wb = new Workbook();
    wb.addWorksheet("Original").getCell("A1").value = 7;
    await expect(
      wb.xlsx.load(bytes, { strict: true, onWarning: (warning) => warnings.push(warning) }),
    ).rejects.toBeInstanceOf(ImportWarningError);
    expect(warnings).toHaveLength(1);
    expect(wb.worksheets[0]?.name).toBe("Original");
    expect(wb.worksheets[0]?.getCell("A1").value).toBe(7);
  });

  it("does not warn for a supported workbook or leak warnings across imports", () => {
    const warnings: ImportWarning[] = [];
    const bytes = fixture();
    expect(() =>
      readWorkbookBuffer(bytes, { strict: true, onWarning: (warning) => warnings.push(warning) }),
    ).not.toThrow();
    expect(warnings).toEqual([]);
  });

  it("reports missing worksheet targets and missing relationship identifiers", () => {
    const missingPart = fixture((parts) => {
      parts.delete("xl/worksheets/sheet1.xml");
    });
    const missingId = fixture((parts) => {
      edit(parts, "xl/workbook.xml", (xml) => xml.replace('r:id="rId1"', 'r:id="unknown"'));
    });
    for (const bytes of [missingPart, missingId]) {
      const warnings: ImportWarning[] = [];
      readWorkbookBuffer(bytes, { onWarning: (warning) => warnings.push(warning) });
      expect(warnings.some((warning) => warning.code === "UNRESOLVED_RELATIONSHIP")).toBe(true);
      expect(() => readWorkbookBuffer(bytes, { strict: true })).toThrow(ImportWarningError);
    }
  });

  it("reports a missing hyperlink relationship and a malformed worksheet root", () => {
    const hyperlink = fixture((parts) =>
      edit(parts, "xl/worksheets/sheet1.xml", (xml) =>
        xml.replace(
          "</worksheet>",
          '<hyperlinks><hyperlink ref="A1" r:id="absent"/></hyperlinks></worksheet>',
        ),
      ),
    );
    const root = fixture((parts) => edit(parts, "xl/worksheets/sheet1.xml", () => "<wrong/>"));
    for (const bytes of [hyperlink, root]) {
      const warnings: ImportWarning[] = [];
      readWorkbookBuffer(bytes, { onWarning: (warning) => warnings.push(warning) });
      expect(warnings.length).toBeGreaterThan(0);
      expect(() => readWorkbookBuffer(bytes, { strict: true })).toThrow(ImportWarningError);
    }
  });

  it("reports foreign extensions and page settings", () => {
    const bytes = fixture((parts) =>
      edit(parts, "xl/worksheets/sheet1.xml", (xml) =>
        xml.replace(
          "</worksheet>",
          '<foreign:feature xmlns:foreign="urn:extension"/><pageSetup orientation="landscape"/></worksheet>',
        ),
      ),
    );
    const warnings: ImportWarning[] = [];
    readWorkbookBuffer(bytes, { onWarning: (warning) => warnings.push(warning) });
    expect(warnings.map((warning) => warning.feature)).toEqual(
      expect.arrayContaining(["XML namespace", "pageSetup"]),
    );
  });

  it("lets observer errors propagate and validates options", () => {
    const bytes = fixture((parts) => {
      parts.set("xl/tables/table1.xml", new TextEncoder().encode("<table/>"));
    });
    const error = new Error("observer failed");
    expect(() =>
      readWorkbookBuffer(bytes, {
        onWarning: () => {
          throw error;
        },
      }),
    ).toThrow(error);
    expect(() => readWorkbookBuffer(bytes, { strict: "yes" as never })).toThrow(TypeError);
    expect(() => readWorkbookBuffer(bytes, { onWarning: true as never })).toThrow(TypeError);
  });
});
