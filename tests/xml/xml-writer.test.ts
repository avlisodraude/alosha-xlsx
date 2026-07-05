import { describe, expect, it } from "vitest";
import { escapeXmlAttr, escapeXmlText, XmlWriter } from "../../src/xml/index.js";

describe("escapeXmlText", () => {
  it("escapes & < > (ampersand first)", () => {
    expect(escapeXmlText("a<b>&c")).toBe("a&lt;b&gt;&amp;c");
    expect(escapeXmlText("<&>")).toBe("&lt;&amp;&gt;");
  });

  it("leaves quotes untouched (text context)", () => {
    expect(escapeXmlText(`he said "hi" don't`)).toBe(`he said "hi" don't`);
  });
});

describe("escapeXmlAttr", () => {
  it("escapes the text set plus double and single quotes", () => {
    expect(escapeXmlAttr(`a"b'c<d>&`)).toBe("a&quot;b&apos;c&lt;d&gt;&amp;");
  });
});

describe("XmlWriter", () => {
  it("emits the XML declaration by default", () => {
    expect(new XmlWriter().toString()).toBe(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    );
  });

  it("omits the declaration when disabled", () => {
    expect(new XmlWriter({ declaration: false }).toString()).toBe("");
  });

  it("builds nested elements via open/el/close", () => {
    const xml = new XmlWriter({ declaration: false })
      .open("si")
      .el("t", "hi")
      .close("si")
      .toString();
    expect(xml).toBe("<si><t>hi</t></si>");
  });

  it("renders leaf elements self-closing", () => {
    expect(new XmlWriter({ declaration: false }).leaf("col", { min: 1, max: 1 }).toString()).toBe(
      '<col min="1" max="1"/>',
    );
  });

  it("omits undefined and null attributes", () => {
    const xml = new XmlWriter({ declaration: false })
      .leaf("c", { r: "A1", t: undefined, s: null })
      .toString();
    expect(xml).toBe('<c r="A1"/>');
  });

  it("renders booleans as 1/0", () => {
    const xml = new XmlWriter({ declaration: false })
      .leaf("col", { customWidth: true, hidden: false })
      .toString();
    expect(xml).toBe('<col customWidth="1" hidden="0"/>');
  });

  it("escapes attribute values", () => {
    const xml = new XmlWriter({ declaration: false }).leaf("a", { v: `x"<&` }).toString();
    expect(xml).toBe('<a v="x&quot;&lt;&amp;"/>');
  });

  it("escapes element text content", () => {
    const xml = new XmlWriter({ declaration: false }).el("t", "a<b&c").toString();
    expect(xml).toBe("<t>a&lt;b&amp;c</t>");
  });

  it('adds xml:space="preserve" to <t> with leading/trailing whitespace', () => {
    expect(new XmlWriter({ declaration: false }).el("t", " hi ").toString()).toBe(
      '<t xml:space="preserve"> hi </t>',
    );
    expect(new XmlWriter({ declaration: false }).el("t", "hi").toString()).toBe("<t>hi</t>");
  });

  it("does not add xml:space to non-<t> elements", () => {
    expect(new XmlWriter({ declaration: false }).el("v", " x ").toString()).toBe("<v> x </v>");
  });

  it("appends raw markup verbatim", () => {
    expect(new XmlWriter({ declaration: false }).raw("<rPr><b/></rPr>").toString()).toBe(
      "<rPr><b/></rPr>",
    );
  });

  it("toBytes round-trips to toString", () => {
    const w = new XmlWriter().open("sst").el("t", "café & tea").close("sst");
    expect(new TextDecoder().decode(w.toBytes())).toBe(w.toString());
  });
});
