import { describe, expect, it } from "vitest";
import type { RichText } from "../../../src/model/index.js";
import { parseSharedStrings } from "../../../src/xlsx/read/shared-strings.js";
import { createSharedStrings } from "../../../src/xlsx/shared-strings.js";

describe("parseSharedStrings", () => {
  it("returns [] for an empty/absent part", () => {
    expect(parseSharedStrings("")).toEqual([]);
  });

  it("returns [] for a self-closing or childless <sst>", () => {
    expect(
      parseSharedStrings(
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
      ),
    ).toEqual([]);
  });

  it("parses plain <si><t> entries in order", () => {
    const xml =
      '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">' +
      "<si><t>a</t></si><si><t>b</t></si></sst>";
    expect(parseSharedStrings(xml)).toEqual(["a", "b"]);
  });

  it("preserves significant whitespace tagged with xml:space=preserve", () => {
    const xml = '<sst xmlns="ns"><si><t xml:space="preserve"> x </t></si></sst>';
    expect(parseSharedStrings(xml)).toEqual([" x "]);
  });

  it("decodes escaped entities in text content", () => {
    const xml = '<sst xmlns="ns"><si><t>a&lt;b&gt;&amp;c</t></si></sst>';
    expect(parseSharedStrings(xml)).toEqual(["a<b>&c"]);
  });

  it("treats a self-closing <t> as an empty string", () => {
    const xml = '<sst xmlns="ns"><si><t/></si></sst>';
    expect(parseSharedStrings(xml)).toEqual([""]);
  });

  it("parses <r> runs into a RichText, with a plain run carrying no font", () => {
    const xml =
      '<sst xmlns="ns"><si><r><rPr><b/><color rgb="FFFF0000"/></rPr><t>bold</t></r>' +
      '<r><t xml:space="preserve"> plain</t></r></si></sst>';
    const [entry] = parseSharedStrings(xml);
    expect(entry).toEqual({
      richText: [
        { text: "bold", font: { bold: true, color: { argb: "FFFF0000" } } },
        { text: " plain" },
      ],
    });
  });

  it("inverts renderRunProps: name, size, italic, underline, scheme", () => {
    const xml =
      '<sst xmlns="ns"><si><r><rPr><rFont val="Arial"/><i/><sz val="12"/>' +
      '<u val="double"/></rPr><t>styled</t></r></si></sst>';
    const [entry] = parseSharedStrings(xml) as [RichText];
    expect(entry.richText[0]).toEqual({
      text: "styled",
      font: { name: "Arial", italic: true, size: 12, underline: "double" },
    });
  });

  it("reads a bare <u/> as underline: true", () => {
    const xml = '<sst xmlns="ns"><si><r><rPr><u/></rPr><t>x</t></r></si></sst>';
    const [entry] = parseSharedStrings(xml) as [RichText];
    expect(entry.richText[0]?.font?.underline).toBe(true);
  });

  it("reads theme and indexed colors", () => {
    const xml =
      '<sst xmlns="ns"><si><r><rPr><color theme="4"/></rPr><t>a</t></r></si>' +
      '<si><r><rPr><color indexed="8"/></rPr><t>b</t></r></si></sst>';
    const [a, b] = parseSharedStrings(xml) as [RichText, RichText];
    expect(a.richText[0]?.font?.color).toEqual({ theme: 4 });
    expect(b.richText[0]?.font?.color).toEqual({ indexed: 8 });
  });

  it("reads vertAlign, scheme, family, charset, strike, outline", () => {
    const xml =
      '<sst xmlns="ns"><si><r><rPr><vertAlign val="superscript"/><scheme val="minor"/>' +
      '<family val="2"/><charset val="1"/><strike/><outline/></rPr><t>x</t></r></si></sst>';
    const [entry] = parseSharedStrings(xml) as [RichText];
    expect(entry.richText[0]?.font).toEqual({
      vertAlign: "superscript",
      scheme: "minor",
      family: 2,
      charset: 1,
      strike: true,
      outline: true,
    });
  });

  it("round-trips the writer's own output for plain strings and rich text", () => {
    const sst = createSharedStrings();
    sst.add("a");
    sst.add(" x ");
    sst.add({
      richText: [
        { text: "hello ", font: { bold: true } },
        { text: "world", font: { italic: true, color: { argb: "FF00FF00" } } },
      ],
    });
    const xml = sst.toXml() ?? "";
    expect(parseSharedStrings(xml)).toEqual([
      "a",
      " x ",
      {
        richText: [
          { text: "hello ", font: { bold: true } },
          { text: "world", font: { italic: true, color: { argb: "FF00FF00" } } },
        ],
      },
    ]);
  });
});
