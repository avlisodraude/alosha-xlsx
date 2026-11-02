import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";
import { readWorkbookBuffer, writeWorkbookBuffer } from "../../src/index.js";
import { NS } from "../../src/xlsx/parts/constants.js";
import { openPackage } from "../../src/xlsx/read/package.js";
import { readZipArchive, writeZipArchive } from "../../src/zip-layer.js";

// Deliberately handwritten OPC package: it does not rely on this library's XML writer.
function independentPackage(sheet: string, rootNamespace: string = NS.packageRels): Uint8Array {
  const office = NS.relationships;
  const parts = {
    "[Content_Types].xml": `<c:Types xmlns:c="${NS.contentTypes}"><c:Override PartName="/xl/book.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></c:Types>`,
    "_rels/.rels": `<z:Relationships xmlns:z="${rootNamespace}"><z:Relationship Id="doc" Type="${office}/officeDocument" Target="xl/book.xml"/></z:Relationships>`,
    "xl/book.xml": `<s:workbook xmlns:s="${NS.spreadsheetml}" xmlns:q="${office}"><s:sheets><s:sheet name="O&apos;Brien" sheetId="1" q:id="sheet"/></s:sheets></s:workbook>`,
    "xl/_rels/book.xml.rels": `<p:Relationships xmlns:p="${NS.packageRels}"><p:Relationship Id="sheet" Type="${office}/worksheet" Target="s.xml"/></p:Relationships>`,
    "xl/s.xml": sheet,
  };
  return writeZipArchive(
    new Map(Object.entries(parts).map(([key, value]) => [key, new TextEncoder().encode(value)])),
  );
}

describe("namespace-aware imports", () => {
  it("reads independently authored prefixed package wiring and worksheet content", () => {
    const bytes = independentPackage(
      `<m:worksheet xmlns:m="${NS.spreadsheetml}"><m:sheetData><m:row r="1"><m:c r="A1" t="inlineStr"><m:is><m:t>Hello</m:t></m:is></m:c></m:row></m:sheetData></m:worksheet>`,
    );
    expect(readWorkbookBuffer(bytes).getWorksheet("O'Brien")?.getCell("A1").value).toBe("Hello");
  });

  it("skips foreign subtrees and restores prefix bindings after nested redeclarations", () => {
    const bytes = independentPackage(`<s:worksheet xmlns:s="${NS.spreadsheetml}">
      <s:sheetData xmlns:s="urn:foreign"><real:sheetData xmlns:real="${NS.spreadsheetml}"><real:row r="1"><real:c r="A1"><real:v>999</real:v></real:c></real:row></real:sheetData></s:sheetData>
      <s:sheetData><s:row r="2"><s:c r="A2"><s:v>42</s:v></s:c></s:row></s:sheetData>
    </s:worksheet>`);
    const ws = readWorkbookBuffer(bytes).worksheets[0];
    expect(ws?.findCell("A1")).toBeUndefined();
    expect(ws?.getCell("A2").value).toBe(42);
  });

  it("ignores elements that reset the default namespace", () => {
    const bytes = independentPackage(
      `<worksheet xmlns="${NS.spreadsheetml}"><sheetData xmlns=""><row r="1"><c r="A1"><v>99</v></c></row></sheetData><sheetData><row r="2"><c r="A2"><v>7</v></c></row></sheetData></worksheet>`,
    );
    const ws = readWorkbookBuffer(bytes).worksheets[0];
    expect(ws?.findCell("A1")).toBeUndefined();
    expect(ws?.getCell("A2").value).toBe(7);
  });

  it("rejects package roots bound to a foreign URI despite familiar names", () => {
    expect(() => openPackage(independentPackage("<worksheet/>", "urn:foreign"))).toThrow(
      /officeDocument/,
    );
  });

  it.each([
    "urn:foreign",
    "toString",
    "constructor",
  ])("does not accept relationship id attributes in %s", (uri) => {
    const bytes = independentPackage(
      `<worksheet xmlns="${NS.spreadsheetml}"><sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData></worksheet>`,
    );
    const parts = readZipArchive(bytes);
    const xml = new TextDecoder()
      .decode(parts.get("xl/book.xml"))
      .replace(`xmlns:q="${NS.relationships}"`, `xmlns:q="${uri}"`);
    parts.set("xl/book.xml", new TextEncoder().encode(xml));
    expect(readWorkbookBuffer(writeZipArchive(parts)).worksheets[0]?.rowCount).toBe(0);
  });

  it("loads an ExcelJS-produced workbook after systematic namespace prefix changes", async () => {
    const wb = new ExcelJS.Workbook();
    wb.creator = "Independent producer";
    const ws = wb.addWorksheet("Foreign");
    ws.getCell("A1").value = { richText: [{ text: "Bold", font: { bold: true } }] };
    ws.getCell("B1").value = 12.5;
    ws.getCell("B1").numFmt = "0.00";
    ws.getCell("C1").value = { text: "Link", hyperlink: "https://example.com" };
    ws.getCell("A1").note = "A note";
    const png = Buffer.from(
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/b9sAAAAASUVORK5CYII=",
      "base64",
    );
    const image = wb.addImage({ buffer: png, extension: "png" } as unknown as Parameters<
      typeof wb.addImage
    >[0]);
    ws.addImage(image, "D1:E2");
    const parts = readZipArchive(new Uint8Array(await wb.xlsx.writeBuffer()));
    for (const [path, data] of parts) {
      if (!path.endsWith(".xml") && !path.endsWith(".rels")) continue;
      let xml = new TextDecoder().decode(data);
      const prefixes = [...xml.matchAll(/xmlns:([A-Za-z_][\w.-]*)="[^"]*"/g)].map(
        (match) => match[1],
      );
      for (const [i, prefix] of prefixes.entries()) {
        xml = xml
          .replace(new RegExp(`\\b${prefix}:`, "g"), `renamed${i}:`)
          .replace(`xmlns:${prefix}=`, `xmlns:renamed${i}=`);
      }
      if (/xmlns="[^"]+"/.test(xml)) {
        xml = xml
          .replace('xmlns="', 'xmlns:defaultPart="')
          .replace(/<(\/?)([A-Za-z_][\w.-]*)(?=[\s/>])/g, "<$1defaultPart:$2");
      }
      parts.set(path, new TextEncoder().encode(xml));
    }
    const loaded = readWorkbookBuffer(writeZipArchive(parts));
    const sheet = loaded.getWorksheet("Foreign");
    expect(loaded.creator).toBe("Independent producer");
    expect(sheet?.getCell("A1").value).toEqual({
      richText: [{ text: "Bold", font: { bold: true } }],
    });
    expect(sheet?.getCell("B1").numFmt).toBe("0.00");
    expect(sheet?.getCell("C1").value).toMatchObject({ hyperlink: "https://example.com" });
    expect(sheet?.getCell("A1").note).toBeDefined();
    expect(sheet?.images).toHaveLength(1);
    expect(loaded.getImage(0)?.buffer).toEqual(new Uint8Array(png));
    const round = readWorkbookBuffer(writeWorkbookBuffer(loaded));
    expect(round.getWorksheet("Foreign")?.getCell("B1").value).toBe(12.5);
  });
});
