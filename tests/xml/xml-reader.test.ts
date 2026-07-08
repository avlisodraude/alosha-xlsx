import { describe, expect, it } from "vitest";
import { cursor, localName, parseXml, type XmlEvent } from "../../src/xml/index.js";

function events(xml: string): XmlEvent[] {
  return [...parseXml(xml)];
}

describe("parseXml", () => {
  it("skips the <?xml?> declaration", () => {
    expect(events('<?xml version="1.0" encoding="UTF-8"?><a/>')).toEqual([
      { kind: "open", name: "a", attrs: {}, selfClosing: true },
    ]);
  });

  it("skips comments", () => {
    expect(events("<a><!-- a comment --></a>")).toEqual([
      { kind: "open", name: "a", attrs: {}, selfClosing: false },
      { kind: "close", name: "a" },
    ]);
  });

  it("skips a leading DOCTYPE-shaped declaration", () => {
    expect(events('<!DOCTYPE foo><a attr="1"/>')).toEqual([
      { kind: "open", name: "a", attrs: { attr: "1" }, selfClosing: true },
    ]);
  });

  it("parses nested elements with open/close pairs", () => {
    expect(events("<a><b><c/></b></a>")).toEqual([
      { kind: "open", name: "a", attrs: {}, selfClosing: false },
      { kind: "open", name: "b", attrs: {}, selfClosing: false },
      { kind: "open", name: "c", attrs: {}, selfClosing: true },
      { kind: "close", name: "b" },
      { kind: "close", name: "a" },
    ]);
  });

  it("handles self-closing tags with and without a space before the slash", () => {
    expect(events('<a x="1"/><b/>')).toEqual([
      { kind: "open", name: "a", attrs: { x: "1" }, selfClosing: true },
      { kind: "open", name: "b", attrs: {}, selfClosing: true },
    ]);
    expect(events('<a x="1" />')).toEqual([
      { kind: "open", name: "a", attrs: { x: "1" }, selfClosing: true },
    ]);
  });

  it("parses single- and double-quoted attributes, including several on one tag", () => {
    expect(events(`<a x="1" y='2' z="three"/>`)).toEqual([
      { kind: "open", name: "a", attrs: { x: "1", y: "2", z: "three" }, selfClosing: true },
    ]);
  });

  it("tolerates a literal > inside a quoted attribute value", () => {
    expect(events('<a b="1>2"/>')).toEqual([
      { kind: "open", name: "a", attrs: { b: "1>2" }, selfClosing: true },
    ]);
  });

  it("keeps namespace prefixes verbatim on tag and attribute names", () => {
    expect(
      events('<cp:coreProperties xmlns:cp="ns"><r:id val="rId1"/></cp:coreProperties>'),
    ).toEqual([
      { kind: "open", name: "cp:coreProperties", attrs: { "xmlns:cp": "ns" }, selfClosing: false },
      { kind: "open", name: "r:id", attrs: { val: "rId1" }, selfClosing: true },
      { kind: "close", name: "cp:coreProperties" },
    ]);
  });

  it("decodes the five named entities", () => {
    expect(events("<t>a &amp; b &lt;c&gt; &quot;d&quot; &apos;e&apos;</t>")).toEqual([
      { kind: "open", name: "t", attrs: {}, selfClosing: false },
      { kind: "text", text: `a & b <c> "d" 'e'` },
      { kind: "close", name: "t" },
    ]);
  });

  it("decodes numeric decimal and hex entities", () => {
    expect(events("<t>&#65;&#x42;&#10;</t>")).toEqual([
      { kind: "open", name: "t", attrs: {}, selfClosing: false },
      { kind: "text", text: "AB\n" },
      { kind: "close", name: "t" },
    ]);
  });

  it("decodes entities inside attribute values", () => {
    expect(events('<a b="x &amp; y"/>')).toEqual([
      { kind: "open", name: "a", attrs: { b: "x & y" }, selfClosing: true },
    ]);
  });

  it("leaves unknown entities untouched", () => {
    expect(events("<t>&unknown;</t>")).toEqual([
      { kind: "open", name: "t", attrs: {}, selfClosing: false },
      { kind: "text", text: "&unknown;" },
      { kind: "close", name: "t" },
    ]);
  });

  it("decodes Excel's _xHHHH_ escape token (written by our own xml-writer for XML-illegal chars)", () => {
    expect(events("<t>a_x000B_b</t>")).toEqual([
      { kind: "open", name: "t", attrs: {}, selfClosing: false },
      { kind: "text", text: "a\x0Bb" },
      { kind: "close", name: "t" },
    ]);
  });

  it("decodes the _xHHHH_ escape token inside attribute values", () => {
    expect(events('<a b="x_x000B_y"/>')).toEqual([
      { kind: "open", name: "a", attrs: { b: "x\x0By" }, selfClosing: true },
    ]);
  });

  it("emits CDATA content as literal text (no entity decoding)", () => {
    expect(events("<t><![CDATA[a & b <c>]]></t>")).toEqual([
      { kind: "open", name: "t", attrs: {}, selfClosing: false },
      { kind: "text", text: "a & b <c>" },
      { kind: "close", name: "t" },
    ]);
  });

  it("preserves significant leading/trailing whitespace in text nodes", () => {
    expect(events('<t xml:space="preserve">  hi  </t>')).toEqual([
      { kind: "open", name: "t", attrs: { "xml:space": "preserve" }, selfClosing: false },
      { kind: "text", text: "  hi  " },
      { kind: "close", name: "t" },
    ]);
  });

  it("handles an element with no attributes and mixed text/child content", () => {
    expect(events("<a>before<b/>after</a>")).toEqual([
      { kind: "open", name: "a", attrs: {}, selfClosing: false },
      { kind: "text", text: "before" },
      { kind: "open", name: "b", attrs: {}, selfClosing: true },
      { kind: "text", text: "after" },
      { kind: "close", name: "a" },
    ]);
  });
});

describe("localName", () => {
  it("strips a namespace prefix", () => {
    expect(localName("cp:coreProperties")).toBe("coreProperties");
    expect(localName("r:id")).toBe("id");
  });

  it("returns the name unchanged when there is no prefix", () => {
    expect(localName("worksheet")).toBe("worksheet");
  });
});

describe("cursor", () => {
  it("next() and peek() walk the same stream without consuming on peek", () => {
    const c = cursor("<a><b/></a>");
    expect(c.peek()).toEqual({ kind: "open", name: "a", attrs: {}, selfClosing: false });
    expect(c.peek()).toEqual({ kind: "open", name: "a", attrs: {}, selfClosing: false });
    expect(c.next()).toEqual({ kind: "open", name: "a", attrs: {}, selfClosing: false });
    expect(c.next()).toEqual({ kind: "open", name: "b", attrs: {}, selfClosing: true });
    expect(c.next()).toEqual({ kind: "close", name: "a" });
    expect(c.next()).toBeUndefined();
  });

  it("textUntilClose collects raw text up to the matching close", () => {
    const c = cursor("<t>hello world</t><next/>");
    expect(c.next()).toEqual({ kind: "open", name: "t", attrs: {}, selfClosing: false });
    expect(c.textUntilClose("t")).toBe("hello world");
    expect(c.next()).toEqual({ kind: "open", name: "next", attrs: {}, selfClosing: true });
  });

  it("textUntilClose ignores markup from other elements while accumulating text", () => {
    const c = cursor("<r><rPr><b/></rPr><t>hi</t></r>");
    expect(c.next()).toEqual({ kind: "open", name: "r", attrs: {}, selfClosing: false });
    expect(c.textUntilClose("r")).toBe("hi");
  });

  it("textUntilClose tracks nested same-name depth", () => {
    const c = cursor("<a>outer<a>inner</a>tail</a><after/>");
    expect(c.next()).toEqual({ kind: "open", name: "a", attrs: {}, selfClosing: false });
    expect(c.textUntilClose("a")).toBe("outerinnertail");
    expect(c.next()).toEqual({ kind: "open", name: "after", attrs: {}, selfClosing: true });
  });

  it("eachChild invokes onOpen for every direct child then consumes the container's close", () => {
    const c = cursor("<row><c r='A1'/><c r='B1'/></row><next/>");
    expect(c.next()).toEqual({ kind: "open", name: "row", attrs: {}, selfClosing: false });
    const seen: string[] = [];
    c.eachChild("row", (ev) => seen.push(ev.attrs.r ?? ""));
    expect(seen).toEqual(["A1", "B1"]);
    expect(c.next()).toEqual({ kind: "open", name: "next", attrs: {}, selfClosing: true });
  });

  it("eachChild lets onOpen recurse into grandchildren via the shared cursor", () => {
    const c = cursor("<sheetData><row><c r='A1'/><c r='A2'/></row></sheetData>");
    expect(c.next()).toEqual({ kind: "open", name: "sheetData", attrs: {}, selfClosing: false });
    const cells: string[] = [];
    c.eachChild("sheetData", (rowEv) => {
      expect(rowEv.name).toBe("row");
      c.eachChild("row", (cellEv) => cells.push(cellEv.attrs.r ?? ""));
    });
    expect(cells).toEqual(["A1", "A2"]);
  });

  it("eachChild on an empty container invokes onOpen zero times", () => {
    const c = cursor("<row></row>");
    expect(c.next()).toEqual({ kind: "open", name: "row", attrs: {}, selfClosing: false });
    let calls = 0;
    c.eachChild("row", () => calls++);
    expect(calls).toBe(0);
  });
});
