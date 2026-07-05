// Workbook: the root aggregate. Sheets are held in an insertion-ordered array plus a `Map<id,
// Worksheet>` index, so `worksheets` is O(n) with no 1-based sparse hole (docs/API-AUDIT.md §5).
// `orderNo` is kept equal to the array index; the deprecated tabColor-as-string branch is dropped.

import { DuplicateWorksheetNameError } from "../errors.js";
import {
  type AddWorksheetOptions,
  validateWorksheetNameFormat,
  Worksheet,
  type WorksheetModel,
} from "./worksheet.js";

/** Workbook-wide flags — notably the 1904 date system. */
export interface WorkbookProperties {
  date1904?: boolean;
  [key: string]: unknown;
}

/** Calculation chain flags (e.g. `fullCalcOnLoad`). */
export interface CalculationProperties {
  fullCalcOnLoad?: boolean;
  [key: string]: unknown;
}

/** Window layout / active-tab state — stored as plain data (round-trips through the model). */
export interface WorkbookView {
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  activeTab?: number;
  visibility?: string;
  [key: string]: unknown;
}

/** The serializable snapshot of a workbook. */
export interface WorkbookModel {
  creator?: string;
  lastModifiedBy?: string;
  created: Date;
  modified: Date;
  lastPrinted?: Date;
  title?: string;
  subject?: string;
  keywords?: string;
  category?: string;
  description?: string;
  company?: string;
  manager?: string;
  properties: WorkbookProperties;
  calcProperties: CalculationProperties;
  views: WorkbookView[];
  sheets: WorksheetModel[];
  nextId: number;
}

export class Workbook {
  creator?: string;
  lastModifiedBy?: string;
  created: Date;
  modified: Date;
  lastPrinted?: Date;
  title?: string;
  subject?: string;
  keywords?: string;
  category?: string;
  description?: string;
  company?: string;
  manager?: string;

  properties: WorkbookProperties = {};
  calcProperties: CalculationProperties = {};
  views: WorkbookView[] = [];

  private readonly _byId = new Map<number, Worksheet>();
  private _sheets: Worksheet[] = [];
  private _nextId = 1;

  constructor() {
    const now = new Date();
    this.created = now;
    this.modified = now;
  }

  /** Ordered, O(n). A fresh copy so callers cannot mutate the backing order. */
  get worksheets(): Worksheet[] {
    return [...this._sheets];
  }

  addWorksheet(name?: string, options?: Partial<AddWorksheetOptions>): Worksheet {
    const resolvedName = name ?? this.uniqueSheetName();
    validateWorksheetNameFormat(resolvedName);
    const lower = resolvedName.toLowerCase();
    for (const ws of this._sheets) {
      if (ws.name.toLowerCase() === lower) throw new DuplicateWorksheetNameError(resolvedName);
    }

    const id = this._nextId++;
    const ws = new Worksheet(this, id, resolvedName, this._sheets.length);

    if (options) {
      if (options.state) ws.state = options.state;
      if (options.properties) ws.properties = { ...ws.properties, ...options.properties };
      if (options.views) ws.views = options.views;
      if (options.pageSetup !== undefined) ws.pageSetup = options.pageSetup;
      if (options.headerFooter !== undefined) ws.headerFooter = options.headerFooter;
    }

    this._byId.set(id, ws);
    this._sheets.push(ws);
    return ws;
  }

  private uniqueSheetName(): string {
    const taken = new Set(this._sheets.map((w) => w.name.toLowerCase()));
    let n = this._sheets.length + 1;
    let candidate = `Sheet${n}`;
    while (taken.has(candidate.toLowerCase())) {
      n++;
      candidate = `Sheet${n}`;
    }
    return candidate;
  }

  getWorksheet(idOrName?: number | string): Worksheet | undefined {
    if (idOrName === undefined) return this._sheets[0];
    if (typeof idOrName === "number") return this._byId.get(idOrName);
    return this._sheets.find((w) => w.name === idOrName);
  }

  removeWorksheet(idOrName: number | string): void {
    const ws =
      typeof idOrName === "number"
        ? this._byId.get(idOrName)
        : this._sheets.find((w) => w.name === idOrName);
    if (!ws) return;
    this._byId.delete(ws.id);
    this._sheets = this._sheets.filter((w) => w !== ws);
    this.reindex();
  }

  /** Keeps `orderNo` equal to array position after a removal. */
  private reindex(): void {
    this._sheets.forEach((w, i) => {
      w.orderNo = i;
    });
  }

  eachSheet(cb: (ws: Worksheet, id: number) => void): void {
    for (const ws of this._sheets) cb(ws, ws.id);
  }

  // --- model ------------------------------------------------------------------------------------

  get model(): WorkbookModel {
    const m: WorkbookModel = {
      created: this.created,
      modified: this.modified,
      properties: { ...this.properties },
      calcProperties: { ...this.calcProperties },
      views: this.views.map((v) => ({ ...v })),
      sheets: this._sheets.map((w) => w.model),
      nextId: this._nextId,
    };
    if (this.creator !== undefined) m.creator = this.creator;
    if (this.lastModifiedBy !== undefined) m.lastModifiedBy = this.lastModifiedBy;
    if (this.lastPrinted !== undefined) m.lastPrinted = this.lastPrinted;
    if (this.title !== undefined) m.title = this.title;
    if (this.subject !== undefined) m.subject = this.subject;
    if (this.keywords !== undefined) m.keywords = this.keywords;
    if (this.category !== undefined) m.category = this.category;
    if (this.description !== undefined) m.description = this.description;
    if (this.company !== undefined) m.company = this.company;
    if (this.manager !== undefined) m.manager = this.manager;
    return m;
  }

  set model(m: WorkbookModel) {
    this.creator = m.creator;
    this.lastModifiedBy = m.lastModifiedBy;
    this.created = m.created;
    this.modified = m.modified;
    this.lastPrinted = m.lastPrinted;
    this.title = m.title;
    this.subject = m.subject;
    this.keywords = m.keywords;
    this.category = m.category;
    this.description = m.description;
    this.company = m.company;
    this.manager = m.manager;
    this.properties = { ...m.properties };
    this.calcProperties = { ...m.calcProperties };
    this.views = m.views.map((v) => ({ ...v }));

    this._byId.clear();
    this._sheets = [];
    for (const sheetModel of m.sheets) {
      const ws = new Worksheet(this, sheetModel.id, sheetModel.name, sheetModel.orderNo);
      ws.model = sheetModel;
      this._byId.set(ws.id, ws);
      this._sheets.push(ws);
    }
    this._nextId = m.nextId;
  }
}
