// Track C session 2 (docs/COMMENTS-SPEC.md §2): inverting the C1 comments.xml writer. Golden
// inversions for the string/rich/multi-author fragments tests/xlsx/parts/comments.test.ts asserts on
// the write side, plus malformed-part tolerance (never throws).

import { describe, expect, it } from "vitest";
import { parseComments } from "../../../src/xlsx/read/comments.js";

describe("parseComments — golden inversions", () => {
  it("inverts a lone <t> into a plain string, resolving authorId 0 through <authors>", () => {
    const xml =
      '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      "<authors><author>Author</author></authors>" +
      '<commentList><comment ref="B2" authorId="0"><text><t>hello</t></text></comment></commentList>' +
      "</comments>";

    expect(parseComments(xml)).toEqual([{ ref: "B2", texts: "hello", author: "Author" }]);
  });

  it("inverts <r><rPr><t> runs into RichText, matching the writer's rich fragment", () => {
    const xml =
      '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      "<authors><author>Alice</author></authors>" +
      '<commentList><comment ref="C3" authorId="0"><text>' +
      "<r><rPr><b/></rPr><t>bold</t></r>" +
      '<r><t xml:space="preserve"> plain</t></r>' +
      "</text></comment></commentList>" +
      "</comments>";

    expect(parseComments(xml)).toEqual([
      {
        ref: "C3",
        texts: { richText: [{ text: "bold", font: { bold: true } }, { text: " plain" }] },
        author: "Alice",
      },
    ]);
  });

  it("resolves distinct authorIds against a multi-author table", () => {
    const xml =
      '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      "<authors><author>Alice</author><author>Bob</author></authors>" +
      "<commentList>" +
      '<comment ref="A1" authorId="0"><text><t>one</t></text></comment>' +
      '<comment ref="A2" authorId="1"><text><t>two</t></text></comment>' +
      '<comment ref="A3" authorId="0"><text><t>three</t></text></comment>' +
      "</commentList>" +
      "</comments>";

    expect(parseComments(xml)).toEqual([
      { ref: "A1", texts: "one", author: "Alice" },
      { ref: "A2", texts: "two", author: "Bob" },
      { ref: "A3", texts: "three", author: "Alice" },
    ]);
  });
});

describe("parseComments — malformed-part tolerance", () => {
  it("returns no records for a rootless/non-comments document", () => {
    expect(parseComments("<notComments/>")).toEqual([]);
    expect(parseComments("")).toEqual([]);
  });

  it("skips a <comment> with no ref", () => {
    const xml =
      '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      "<authors><author>Author</author></authors>" +
      '<commentList><comment authorId="0"><text><t>orphan</t></text></comment></commentList>' +
      "</comments>";
    expect(parseComments(xml)).toEqual([]);
  });

  it("skips a <comment> whose authorId doesn't resolve against <authors>", () => {
    const xml =
      '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      "<authors><author>Author</author></authors>" +
      '<commentList><comment ref="A1" authorId="5"><text><t>x</t></text></comment></commentList>' +
      "</comments>";
    expect(parseComments(xml)).toEqual([]);
  });

  it("keeps well-formed comments alongside a malformed sibling", () => {
    const xml =
      '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      "<authors><author>Author</author></authors>" +
      "<commentList>" +
      '<comment authorId="0"><text><t>orphan</t></text></comment>' +
      '<comment ref="A1" authorId="0"><text><t>kept</t></text></comment>' +
      "</commentList>" +
      "</comments>";
    expect(parseComments(xml)).toEqual([{ ref: "A1", texts: "kept", author: "Author" }]);
  });
});
