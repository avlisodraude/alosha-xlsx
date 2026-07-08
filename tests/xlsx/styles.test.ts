import { describe, expect, it } from "vitest";
import {
  createStyles,
  renderBorder,
  renderFill,
  renderFont,
  renderNumFmt,
  renderXf,
} from "../../src/xlsx/styles.js";

describe("createStyles — baseline seeds", () => {
  it("seeds the mandatory Excel pools and a default cellXf", () => {
    const xml = createStyles().toXml();
    expect(xml).toContain(
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    );
    // fonts[0] = Calibri 11, theme text colour, minor scheme.
    expect(xml).toContain(
      '<fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>',
    );
    // fills[0] = none, fills[1] = gray125, in that order.
    expect(xml).toContain(
      '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>',
    );
    // borders[0] = fully empty border with all sides present.
    expect(xml).toContain(
      '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>',
    );
    // one default cellStyleXf and one default cellXf.
    expect(xml).toContain(
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>',
    );
    expect(xml).toContain(
      '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>',
    );
    expect(xml).toContain(
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>',
    );
    expect(xml).toContain('<dxfs count="0"/>');
    expect(xml).toContain(
      '<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>',
    );
  });

  it("omits <numFmts> entirely when no custom formats are used", () => {
    expect(createStyles().toXml()).not.toContain("<numFmts");
  });

  it("emits the top-level pools in CT_Stylesheet schema order", () => {
    const styles = createStyles();
    styles.addStyle({ numFmt: "0.000" });
    const xml = styles.toXml();
    const order = [
      "<numFmts",
      "<fonts",
      "<fills",
      "<borders",
      "<cellStyleXfs",
      "<cellXfs",
      "<cellStyles",
      "<dxfs",
      "<tableStyles",
    ];
    const positions = order.map((tag) => xml.indexOf(tag));
    expect(positions.every((p) => p >= 0)).toBe(true);
    expect(positions).toEqual([...positions].sort((a, b) => a - b));
  });
});

describe("addStyle — interning and dedup", () => {
  it("returns 0 for undefined and for a structurally-default style", () => {
    const styles = createStyles();
    expect(styles.addStyle(undefined)).toBe(0);
    expect(styles.addStyle({})).toBe(0);
    expect(styles.addStyle({ numFmt: "General" })).toBe(0); // General → numFmtId 0 → default xf
  });

  it("dedups structurally-identical styles across distinct objects", () => {
    const styles = createStyles();
    const first = styles.addStyle({ font: { bold: true } });
    const second = styles.addStyle({ font: { bold: true } }); // different object, same shape
    expect(second).toBe(first);
    expect(styles.addStyle({ font: { italic: true } })).not.toBe(first);
  });

  it("reuses one fontId when the same font backs several styles", () => {
    const styles = createStyles();
    styles.addStyle({ font: { bold: true } });
    styles.addStyle({ font: { bold: true }, numFmt: "0.00%" }); // same font, different numFmt
    // default + bold only → 2 fonts total.
    expect(styles.toXml()).toContain('<fonts count="2">');
  });

  it("takes the WeakMap fast path for a shared style object", () => {
    const styles = createStyles();
    const shared = {
      fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF0000" } },
    } as const;
    const a = styles.addStyle(shared);
    const b = styles.addStyle(shared);
    expect(b).toBe(a);
    expect(styles.toXml()).toContain('<cellXfs count="2">'); // default + one
  });
});

describe("number formats", () => {
  it("maps built-in codes to reserved ids without allocating a <numFmt>", () => {
    const styles = createStyles();
    styles.addStyle({ numFmt: "0.00%" });
    const xml = styles.toXml();
    expect(xml).not.toContain("<numFmts");
    expect(xml).toContain('numFmtId="10"');
    expect(xml).toContain('applyNumberFormat="1"');
  });

  it("allocates custom codes from 164 upward and dedups by code", () => {
    const styles = createStyles();
    styles.addStyle({ numFmt: "0.000" });
    styles.addStyle({ numFmt: "0.000" }); // same custom code
    styles.addStyle({ numFmt: '"€"#,##0.00' }); // second custom code
    const xml = styles.toXml();
    expect(xml).toContain('<numFmts count="2">');
    expect(xml).toContain('<numFmt numFmtId="164" formatCode="0.000"/>');
    expect(xml).toContain('<numFmt numFmtId="165" formatCode="&quot;€&quot;#,##0.00"/>');
  });

  it("gives each distinct custom dxf number-format a distinct id (not all 164)", () => {
    const styles = createStyles();
    styles.addDxf({ numFmt: "0.000" });
    styles.addDxf({ numFmt: "0.000" }); // same code -> same dxf, deduped in the pool
    styles.addDxf({ numFmt: '"€"#,##0.00' }); // different custom code -> distinct id
    styles.addDxf({ numFmt: "0.00%" }); // built-in code -> keeps its reserved id
    const xml = styles.toXml();
    expect(xml).toContain('<numFmt numFmtId="164" formatCode="0.000"/>');
    expect(xml).toContain('<numFmt numFmtId="165" formatCode="&quot;€&quot;#,##0.00"/>');
    expect(xml).toContain('<numFmt numFmtId="10" formatCode="0.00%"/>'); // built-in id preserved
    // Identical dxfs still dedupe: only one dxf carries formatCode="0.000".
    expect(xml.match(/formatCode="0\.000"/g)?.length).toBe(1);
  });

  it("renderNumFmt escapes the format code", () => {
    expect(renderNumFmt(164, '0.0"<>"')).toBe(
      '<numFmt numFmtId="164" formatCode="0.0&quot;&lt;&gt;&quot;"/>',
    );
  });
});

describe("renderFont", () => {
  it("orders elements per CT_Font and uses <name> not <rFont>", () => {
    expect(
      renderFont({
        bold: true,
        italic: true,
        size: 12,
        color: { argb: "FF112233" },
        name: "Arial",
      }),
    ).toBe('<font><b/><i/><sz val="12"/><color rgb="FF112233"/><name val="Arial"/></font>');
  });

  it("renders underline variants", () => {
    expect(renderFont({ underline: true })).toBe("<font><u/></font>");
    expect(renderFont({ underline: "double" })).toBe('<font><u val="double"/></font>');
    expect(renderFont({ underline: "none" })).toBe("<font></font>");
  });
});

describe("renderFill", () => {
  it("renders a solid pattern fill with foreground and background colours", () => {
    expect(
      renderFill({
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFF0000" },
        bgColor: { indexed: 64 },
      }),
    ).toBe(
      '<fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/><bgColor indexed="64"/></patternFill></fill>',
    );
  });

  it("self-closes a colourless pattern fill", () => {
    expect(renderFill({ type: "pattern", pattern: "none" })).toBe(
      '<fill><patternFill patternType="none"/></fill>',
    );
  });

  it("renders angle and path gradient fills", () => {
    expect(
      renderFill({
        type: "gradient",
        gradient: "angle",
        degree: 90,
        stops: [
          { position: 0, color: { argb: "FFFF0000" } },
          { position: 1, color: { argb: "FF0000FF" } },
        ],
      }),
    ).toBe(
      '<fill><gradientFill degree="90"><stop position="0"><color rgb="FFFF0000"/></stop><stop position="1"><color rgb="FF0000FF"/></stop></gradientFill></fill>',
    );
    expect(
      renderFill({
        type: "gradient",
        gradient: "path",
        center: { left: 0.5, top: 0.5 },
        stops: [{ position: 0, color: { argb: "FFFFFFFF" } }],
      }),
    ).toBe(
      '<fill><gradientFill type="path" left="0.5" right="0.5" top="0.5" bottom="0.5"><stop position="0"><color rgb="FFFFFFFF"/></stop></gradientFill></fill>',
    );
  });
});

describe("renderBorder", () => {
  it("always emits all sides, self-closing empty ones", () => {
    expect(renderBorder(undefined)).toBe(
      "<border><left/><right/><top/><bottom/><diagonal/></border>",
    );
  });

  it("renders styled sides with colours and diagonal direction flags", () => {
    expect(
      renderBorder({
        top: { style: "thin", color: { argb: "FF000000" } },
        diagonal: { style: "thin", up: true },
      }),
    ).toBe(
      '<border diagonalUp="1"><left/><right/><top style="thin"><color rgb="FF000000"/></top><bottom/><diagonal style="thin"/></border>',
    );
  });
});

describe("renderXf", () => {
  it("sets no apply flags and self-closes for an all-default xf", () => {
    expect(renderXf({ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 })).toBe(
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
    );
  });

  it("sets applyX flags for each non-default component", () => {
    expect(renderXf({ numFmtId: 164, fontId: 2, fillId: 3, borderId: 1 })).toBe(
      '<xf numFmtId="164" fontId="2" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>',
    );
  });

  it("emits alignment with rotation, wrap and reading-order conversions", () => {
    expect(
      renderXf({
        numFmtId: 0,
        fontId: 0,
        fillId: 0,
        borderId: 0,
        alignment: {
          horizontal: "center",
          vertical: "middle",
          textRotation: "vertical",
          wrapText: true,
          readingOrder: "rtl",
        },
      }),
    ).toBe(
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="middle" textRotation="255" wrapText="1" readingOrder="2"/></xf>',
    );
  });

  it("converts negative text rotation to Excel's 90+ encoding", () => {
    expect(
      renderXf({
        numFmtId: 0,
        fontId: 0,
        fillId: 0,
        borderId: 0,
        alignment: { textRotation: -45 },
      }),
    ).toContain('textRotation="135"');
  });

  it("emits protection, keeping a meaningful locked=false", () => {
    expect(
      renderXf({
        numFmtId: 0,
        fontId: 0,
        fillId: 0,
        borderId: 0,
        protection: { locked: false, hidden: true },
      }),
    ).toBe(
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1"><protection locked="0" hidden="1"/></xf>',
    );
  });

  it("omits an empty alignment object rather than flagging applyAlignment", () => {
    expect(renderXf({ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, alignment: {} })).toBe(
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
    );
  });
});
