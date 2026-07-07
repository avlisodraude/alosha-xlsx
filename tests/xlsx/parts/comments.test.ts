// Track C session 1 (docs/COMMENTS-SPEC.md §1): the writer side of `cell.note` — the `comments{n}.xml`
// note/author table, the `vmlDrawing{n}.vml` anchor part, the two worksheet rels + content-types, the
// `<legacyDrawing>` last-child position, and a no-comments regression.

import { describe, expect, it } from "vitest";
import { Workbook } from "../../../src/model/index.js";
import { CONTENT_TYPE, REL_TYPE } from "../../../src/xlsx/parts/constants.js";
import { type RenderedWorksheet, renderWorksheet } from "../../../src/xlsx/parts/worksheet.js";
import { createSharedStrings } from "../../../src/xlsx/shared-strings.js";
import { createStyles } from "../../../src/xlsx/styles.js";
import { writeWorkbookBuffer } from "../../../src/xlsx/write-workbook.js";
import { readZipArchive } from "../../../src/zip-layer.js";

/** Render a single-sheet worksheet through the real pools, returning its xml + rels + extraParts. */
function render(build: (ws: ReturnType<Workbook["addWorksheet"]>) => void): RenderedWorksheet {
  const ws = new Workbook().addWorksheet("Data");
  build(ws);
  return renderWorksheet(ws, {
    styles: createStyles(),
    sharedStrings: createSharedStrings(),
    useSharedStrings: true,
    date1904: false,
    sheetNumber: 1,
  });
}

/** The string data of the extra part at `path` (comments/vml are always string parts). */
function partData(rendered: RenderedWorksheet, path: string): string | undefined {
  const part = rendered.extraParts?.find((p) => p.path === path);
  return typeof part?.data === "string" ? part.data : undefined;
}

describe("worksheet writer — comments.xml", () => {
  it("renders a plain-string comment with the default author as a lone <t>", () => {
    const rendered = render((ws) => {
      ws.getCell("B2").note = { texts: "hello" };
    });
    const xml = partData(rendered, "xl/comments1.xml");

    expect(xml).toContain(
      '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    );
    // Author defaults to "Author" when unset (ExcelJS parity).
    expect(xml).toContain("<authors><author>Author</author></authors>");
    expect(xml).toContain(
      '<commentList><comment ref="B2" authorId="0"><text><t>hello</t></text></comment></commentList>',
    );
  });

  it("renders rich-text runs through the shared run renderer (<r><rPr><t>)", () => {
    const rendered = render((ws) => {
      ws.getCell("C3").note = {
        texts: { richText: [{ text: "bold", font: { bold: true } }, { text: " plain" }] },
        author: "Alice",
      };
    });
    const xml = partData(rendered, "xl/comments1.xml");

    expect(xml).toContain("<authors><author>Alice</author></authors>");
    expect(xml).toContain(
      '<comment ref="C3" authorId="0"><text>' +
        "<r><rPr><b/></rPr><t>bold</t></r>" +
        '<r><t xml:space="preserve"> plain</t></r>' +
        "</text></comment>",
    );
  });

  it("builds a distinct authors table in first-seen order and indexes authorId into it", () => {
    const rendered = render((ws) => {
      ws.getCell("A1").note = { texts: "one", author: "Alice" };
      ws.getCell("A2").note = { texts: "two", author: "Bob" };
      ws.getCell("A3").note = { texts: "three", author: "Alice" };
    });
    const xml = partData(rendered, "xl/comments1.xml");

    expect(xml).toContain("<authors><author>Alice</author><author>Bob</author></authors>");
    expect(xml).toContain('<comment ref="A1" authorId="0">');
    expect(xml).toContain('<comment ref="A2" authorId="1">');
    expect(xml).toContain('<comment ref="A3" authorId="0">');
  });
});

describe("worksheet writer — vmlDrawing.vml", () => {
  it("emits one note shape per comment with a 0-based ClientData anchor", () => {
    const rendered = render((ws) => {
      ws.getCell("B2").note = { texts: "hi" };
    });
    const vml = partData(rendered, "xl/drawings/vmlDrawing1.vml");

    // No XML declaration — VML parts are raw <xml> roots.
    expect(vml).not.toContain("<?xml");
    expect(vml).toContain('<v:shape id="_x0000_s1025" type="#_x0000_t202"');
    expect(vml).toContain('<x:ClientData ObjectType="Note">');
    // B2 -> row 2, col 2 -> 0-based Row 1, Column 1.
    expect(vml).toContain("<x:Row>1</x:Row>");
    expect(vml).toContain("<x:Column>1</x:Column>");
    expect(vml).toContain("<x:AutoFill>False</x:AutoFill>");
  });

  it("gives each comment its own shape id", () => {
    const rendered = render((ws) => {
      ws.getCell("A1").note = { texts: "one" };
      ws.getCell("A2").note = { texts: "two" };
    });
    const vml = partData(rendered, "xl/drawings/vmlDrawing1.vml");
    expect(vml).toContain('id="_x0000_s1025"');
    expect(vml).toContain('id="_x0000_s1026"');
  });
});

describe("worksheet writer — comment rels + <legacyDrawing>", () => {
  it("pushes both extra parts with their content-type declarations", () => {
    const rendered = render((ws) => {
      ws.getCell("B2").note = { texts: "hi" };
    });

    const comments = rendered.extraParts?.find((p) => p.path === "xl/comments1.xml");
    expect(comments?.contentType).toBe(CONTENT_TYPE.comments);

    const vml = rendered.extraParts?.find((p) => p.path === "xl/drawings/vmlDrawing1.vml");
    expect(vml?.defaultFor).toEqual({ extension: "vml", contentType: CONTENT_TYPE.vmlDrawing });
  });

  it("allocates a comments rel then a vmlDrawing rel, and points <legacyDrawing> at the VML", () => {
    const rendered = render((ws) => {
      ws.getCell("B2").note = { texts: "hi" };
    });

    expect(rendered.rels).toContain(
      `<Relationship Id="rId1" Type="${REL_TYPE.comments}" Target="../comments1.xml"/>`,
    );
    expect(rendered.rels).toContain(
      `<Relationship Id="rId2" Type="${REL_TYPE.vmlDrawing}" Target="../drawings/vmlDrawing1.vml"/>`,
    );
    // <legacyDrawing> references the vmlDrawing rId (rId2), not the comments part.
    expect(rendered.xml).toContain('<legacyDrawing r:id="rId2"/>');
  });

  it("emits <legacyDrawing> as the last child, after hyperlinks (rels ordered link → comments → vml)", () => {
    const rendered = render((ws) => {
      ws.getCell("A1").value = { text: "site", hyperlink: "https://example.com" };
      ws.getCell("B2").note = { texts: "hi" };
    });

    // Hyperlink takes rId1; comments rId2; vmlDrawing rId3.
    expect(rendered.xml).toContain('<hyperlink ref="A1" r:id="rId1"/>');
    expect(rendered.xml).toContain('<legacyDrawing r:id="rId3"/>');
    // <legacyDrawing> is the final child of <worksheet>, sitting after <hyperlinks>.
    expect(rendered.xml).toMatch(/<\/hyperlinks><legacyDrawing r:id="rId3"\/><\/worksheet>$/);
  });
});

describe("write orchestrator — comments fold into the archive", () => {
  const dec = new TextDecoder();

  it("writes the comments + vml parts and types both in [Content_Types].xml", () => {
    const wb = new Workbook();
    wb.addWorksheet("Data").getCell("B2").note = { texts: "hi", author: "Alice" };

    const files = readZipArchive(writeWorkbookBuffer(wb));

    expect(files.has("xl/comments1.xml")).toBe(true);
    expect(files.has("xl/drawings/vmlDrawing1.vml")).toBe(true);

    const contentTypes = dec.decode(files.get("[Content_Types].xml"));
    // A vml <Default> and a per-part comments <Override>.
    expect(contentTypes).toContain(
      `<Default Extension="vml" ContentType="${CONTENT_TYPE.vmlDrawing}"/>`,
    );
    expect(contentTypes).toContain(
      `<Override PartName="/xl/comments1.xml" ContentType="${CONTENT_TYPE.comments}"/>`,
    );
  });

  it("leaves a no-comments workbook byte-identical to F0 (no parts, no <legacyDrawing>)", () => {
    // Byte-identical guarantee: the comment code path is inert without a note — the render carries
    // no extraParts and the sheet body gains no <legacyDrawing>.
    const rendered = render((ws) => {
      ws.getCell("A1").value = "just text";
    });
    expect(rendered.extraParts).toBeUndefined();
    expect(rendered.xml).not.toContain("legacyDrawing");

    const wb = new Workbook();
    wb.addWorksheet("Data").getCell("A1").value = "just text";
    const files = readZipArchive(writeWorkbookBuffer(wb));
    expect(files.has("xl/comments1.xml")).toBe(false);
    expect([...files.keys()].some((k) => k.startsWith("xl/drawings/"))).toBe(false);
    const contentTypes = dec.decode(files.get("[Content_Types].xml"));
    expect(contentTypes).not.toContain('Extension="vml"');
    expect(contentTypes).not.toContain("/xl/comments");
  });
});
