// Inverse of the Stage 1.3 workbook part writer (see docs/STAGE-2-SPEC.md §7, and
// src/xlsx/parts/workbook.ts for the forward direction). Reads xl/workbook.xml for the sheet list
// (name / sheetId / state / `r:id`), the `date1904` flag, and the active tab. Part *locations* are not
// here — each `r:id` is resolved to a worksheet part through workbook.xml.rels by the orchestrator.

import { cursor, localName, type XmlCursor, type XmlEvent } from "../../xml/index.js";
import { NS } from "../parts/constants.js";

export interface WorkbookHeader {
  sheets: {
    name: string;
    sheetId: number;
    state: "visible" | "hidden" | "veryHidden";
    rId: string;
  }[];
  date1904: boolean;
  activeTab?: number;
}

type OpenEvent = Extract<XmlEvent, { kind: "open" }>;

/** Parse xl/workbook.xml. A malformed/rootless document yields an empty header (no sheets). */
export function parseWorkbook(xml: string): WorkbookHeader {
  const header: WorkbookHeader = { sheets: [], date1904: false };

  const c = cursor(xml, [NS.spreadsheetml]);
  let ev = c.next();
  while (ev && !(ev.kind === "open" && localName(ev.name) === "workbook")) ev = c.next();
  if (ev?.kind !== "open" || ev.selfClosing) return header;

  c.eachChild(ev.name, (child) => {
    switch (localName(child.name)) {
      case "workbookPr":
        if (truthy(child.attrs.date1904)) header.date1904 = true;
        break;
      case "bookViews":
        if (!child.selfClosing) parseBookViews(c, child.name, header);
        break;
      case "sheets":
        if (!child.selfClosing) parseSheets(c, child.name, header);
        break;
      default:
        if (!child.selfClosing) c.textUntilClose(child.name);
    }
  });

  return header;
}

function parseBookViews(c: XmlCursor, container: string, header: WorkbookHeader): void {
  c.eachChild(container, (child) => {
    if (localName(child.name) === "workbookView") {
      const activeTab = child.attrs.activeTab;
      if (activeTab !== undefined) {
        const n = Number(activeTab);
        if (Number.isFinite(n)) header.activeTab = n;
      }
    } else if (!child.selfClosing) {
      c.textUntilClose(child.name);
    }
  });
}

function parseSheets(c: XmlCursor, container: string, header: WorkbookHeader): void {
  c.eachChild(container, (child) => {
    if (localName(child.name) !== "sheet") {
      if (!child.selfClosing) c.textUntilClose(child.name);
      return;
    }
    header.sheets.push({
      name: child.attrs.name ?? "",
      sheetId: Number(child.attrs.sheetId ?? 0),
      state: normalizeState(child.attrs.state),
      rId: relId(child),
    });
  });
}

/** Relationship attributes are canonicalized to r:id by URI. Retain the qualified-name fallback
 *  for namespace-free fragments passed to the low-level part reader. */
function relId(sheet: OpenEvent): string {
  const direct = sheet.attrs["r:id"];
  if (direct !== undefined) return direct;
  for (const key of Object.keys(sheet.attrs)) {
    if (localName(key) === "id") return sheet.attrs[key] ?? "";
  }
  return "";
}

function normalizeState(state: string | undefined): "visible" | "hidden" | "veryHidden" {
  if (state === "hidden") return "hidden";
  if (state === "veryHidden") return "veryHidden";
  return "visible";
}

/** A boolean-ish OOXML attribute: `"1"`/`"true"` mean true; absent or anything else false. */
function truthy(val: string | undefined): boolean {
  return val === "1" || val === "true";
}
