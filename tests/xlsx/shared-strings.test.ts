import { describe, expect, it } from "vitest";
import type { RichText } from "../../src/model/index.js";
import { createSharedStrings } from "../../src/xlsx/shared-strings.js";

describe("createSharedStrings", () => {
  it("starts empty and renders no XML", () => {
    const sst = createSharedStrings();
    expect(sst.count).toBe(0);
    expect(sst.uniqueCount).toBe(0);
    expect(sst.toXml()).toBeUndefined();
  });

  it("dedups plain strings by value, tracking count vs uniqueCount", () => {
    const sst = createSharedStrings();
    expect(sst.add("a")).toBe(0);
    expect(sst.add("a")).toBe(0);
    expect(sst.add("b")).toBe(1);
    expect(sst.count).toBe(3);
    expect(sst.uniqueCount).toBe(2);
  });

  it("renders the sst with count/uniqueCount attributes and <si><t> items", () => {
    const sst = createSharedStrings();
    sst.add("a");
    sst.add("a");
    sst.add("b");
    const xml = sst.toXml();
    expect(xml).toContain('<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"');
    expect(xml).toContain('count="3"');
    expect(xml).toContain('uniqueCount="2"');
    expect(xml).toContain("<si><t>a</t></si>");
    expect(xml).toContain("<si><t>b</t></si>");
  });

  it("preserves significant whitespace on string items", () => {
    const sst = createSharedStrings();
    sst.add(" x ");
    expect(sst.toXml()).toContain('<si><t xml:space="preserve"> x </t></si>');
  });

  it("escapes string content", () => {
    const sst = createSharedStrings();
    sst.add("a<b>&c");
    expect(sst.toXml()).toContain("<si><t>a&lt;b&gt;&amp;c</t></si>");
  });

  it("dedups rich text by a stable serialization of its runs", () => {
    const sst = createSharedStrings();
    const rich1: RichText = {
      richText: [
        { text: "hello ", font: { bold: true } },
        { text: "world", font: { italic: true } },
      ],
    };
    // Structurally identical but a distinct object with reordered font keys.
    const rich2: RichText = {
      richText: [
        { text: "hello ", font: { bold: true } },
        { text: "world", font: { italic: true } },
      ],
    };
    expect(sst.add(rich1)).toBe(0);
    expect(sst.add(rich2)).toBe(0);
    expect(sst.count).toBe(2);
    expect(sst.uniqueCount).toBe(1);
  });

  it("keys rich text distinctly from a plain string of the same text", () => {
    const sst = createSharedStrings();
    const rich: RichText = { richText: [{ text: "hi" }] };
    expect(sst.add("hi")).toBe(0);
    expect(sst.add(rich)).toBe(1);
    expect(sst.uniqueCount).toBe(2);
  });

  it("distinguishes rich text runs that differ only by font", () => {
    const sst = createSharedStrings();
    sst.add({ richText: [{ text: "x", font: { bold: true } }] });
    sst.add({ richText: [{ text: "x", font: { bold: false } }] });
    expect(sst.uniqueCount).toBe(2);
  });

  it("renders rich text as <r><rPr>…</rPr><t>…</t></r> runs in schema order", () => {
    const sst = createSharedStrings();
    sst.add({
      richText: [
        { text: "bold", font: { bold: true, color: { argb: "FFFF0000" } } },
        { text: " plain" },
      ],
    });
    const xml = sst.toXml() ?? "";
    expect(xml).toContain('<r><rPr><b/><color rgb="FFFF0000"/></rPr><t>bold</t></r>');
    // A run with no font has no <rPr> and preserves its leading space.
    expect(xml).toContain('<r><t xml:space="preserve"> plain</t></r>');
  });

  it("renders font run properties: name, size, italic, underline, scheme", () => {
    const sst = createSharedStrings();
    sst.add({
      richText: [
        { text: "styled", font: { name: "Arial", size: 12, italic: true, underline: "double" } },
      ],
    });
    const xml = sst.toXml() ?? "";
    expect(xml).toContain('<rPr><rFont val="Arial"/><i/><sz val="12"/><u val="double"/></rPr>');
  });
});
