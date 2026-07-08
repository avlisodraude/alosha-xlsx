import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { Cell, Column, Range, Row, version, Workbook, Worksheet } from "../src/index.js";

describe("@alosha/xlsx", () => {
  it("exposes the real package.json version (injected at build time)", () => {
    const pkg = JSON.parse(
      readFileSync(fileURLToPath(new URL("../package.json", import.meta.url)), "utf8"),
    ) as { version: string };
    expect(version).toBe(pkg.version);
    expect(version).not.toBe("0.0.0");
  });

  it("re-exports the core model and address surface from the barrel", () => {
    expect(Workbook).toBeTypeOf("function");
    expect(Worksheet).toBeTypeOf("function");
    expect(Row).toBeTypeOf("function");
    expect(Column).toBeTypeOf("function");
    expect(Cell).toBeTypeOf("function");
    expect(Range).toBeTypeOf("function");
  });

  it("builds a workbook through the public entry point", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Data");
    ws.getCell("A1").value = "hello";
    expect(wb.getWorksheet("Data")?.getCell("A1").value).toBe("hello");
  });
});
