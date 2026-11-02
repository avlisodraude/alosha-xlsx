// Parses `docProps/core.xml` + `docProps/app.xml` back onto a `Workbook` (see
// docs/STAGE-2-SPEC.md §7). Inverts `../parts/core-props.ts` / `../parts/app-props.ts`: unset
// elements leave the model's defaults untouched.

import type { Workbook } from "../../model/index.js";
import { cursor, localName } from "../../xml/index.js";
import { NS } from "../parts/constants.js";

/** Parses a W3CDTF timestamp (`YYYY-MM-DDTHH:mm:ssZ`, as emitted by `toW3CDTF`) into a `Date`. */
export function fromW3CDTF(text: string): Date {
  return new Date(text);
}

/** Applies `docProps/core.xml` (Dublin Core metadata, W3CDTF dates) onto `wb`. */
export function applyCoreProps(wb: Workbook, xml: string): void {
  const c = cursor(xml, [NS.coreProps, NS.dc, NS.dcterms]);
  let ev = c.next();
  while (ev && !(ev.kind === "open" && localName(ev.name) === "coreProperties")) ev = c.next();
  if (ev?.kind !== "open" || ev.selfClosing) return;

  c.eachChild(ev.name, (child) => {
    const name = localName(child.name);
    const text = child.selfClosing ? "" : c.textUntilClose(child.name);
    switch (name) {
      case "creator":
        wb.creator = text;
        break;
      case "title":
        wb.title = text;
        break;
      case "subject":
        wb.subject = text;
        break;
      case "description":
        wb.description = text;
        break;
      case "keywords":
        wb.keywords = text;
        break;
      case "category":
        wb.category = text;
        break;
      case "lastModifiedBy":
        wb.lastModifiedBy = text;
        break;
      case "created":
        wb.created = fromW3CDTF(text);
        break;
      case "modified":
        wb.modified = fromW3CDTF(text);
        break;
    }
  });
}

/** Applies `docProps/app.xml` (Company/Manager) onto `wb`. */
export function applyAppProps(wb: Workbook, xml: string): void {
  const c = cursor(xml, [NS.extProps, NS.docPropsVt]);
  let ev = c.next();
  while (ev && !(ev.kind === "open" && localName(ev.name) === "Properties")) ev = c.next();
  if (ev?.kind !== "open" || ev.selfClosing) return;

  c.eachChild(ev.name, (child) => {
    const name = localName(child.name);
    const text = child.selfClosing ? "" : c.textUntilClose(child.name);
    if (name === "Company") wb.company = text;
    else if (name === "Manager") wb.manager = text;
  });
}
