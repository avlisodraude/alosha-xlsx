// A minimal WorksheetHost double for Cell/Row/Column tests — auto-creates columns/rows on demand,
// like a real Worksheet would, without depending on the (not-yet-built) Worksheet class.

import { colLetterToNumber } from "../../../src/address/index.js";
import { Column } from "../../../src/model/column.js";
import { DataValidations } from "../../../src/model/data-validation.js";
import { Row } from "../../../src/model/row.js";
import type { WorksheetHost } from "../../../src/model/worksheet-host.js";

type EachRowCb = (row: Row, n: number) => void;

export class FakeWorksheetHost implements WorksheetHost {
  readonly properties: { outlineLevelCol: number; outlineLevelRow: number } & Record<
    string,
    unknown
  > = { outlineLevelCol: 0, outlineLevelRow: 0 };

  readonly dataValidations = new DataValidations();

  private readonly columnsByNumber = new Map<number, Column>();
  private readonly columnsByKey = new Map<string, Column>();
  private readonly rows = new Map<number, Row>();

  getColumn(col: number | string): Column {
    if (typeof col === "number") return this.getOrCreateColumn(col);
    const byKey = this.columnsByKey.get(col);
    if (byKey) return byKey;
    return this.getOrCreateColumn(colLetterToNumber(col));
  }

  private getOrCreateColumn(number: number): Column {
    let column = this.columnsByNumber.get(number);
    if (!column) {
      column = new Column(this, number);
      this.columnsByNumber.set(number, column);
    }
    return column;
  }

  /** Registers a column under a lookup key (mirrors Worksheet.setColumnKey, session 4). */
  registerColumnKey(key: string, column: Column): void {
    this.columnsByKey.set(key, column);
  }

  getColumnKey(key: string): Column | undefined {
    return this.columnsByKey.get(key);
  }

  getRow(row: number): Row {
    let r = this.rows.get(row);
    if (!r) {
      r = new Row(this, row);
      this.rows.set(row, r);
    }
    return r;
  }

  eachRow(cb: EachRowCb): void;
  eachRow(opts: { includeEmpty: boolean }, cb: EachRowCb): void;
  eachRow(optsOrCb: { includeEmpty: boolean } | EachRowCb, maybeCb?: EachRowCb): void {
    const includeEmpty = typeof optsOrCb === "object" ? optsOrCb.includeEmpty : false;
    const cb = typeof optsOrCb === "function" ? optsOrCb : (maybeCb as EachRowCb);

    if (includeEmpty) {
      const max = this.maxRowNumber();
      for (let n = 1; n <= max; n++) {
        cb(this.getRow(n), n);
      }
    } else {
      const sortedRowNumbers = [...this.rows.keys()].sort((a, b) => a - b);
      for (const n of sortedRowNumbers) {
        cb(this.getRow(n), n);
      }
    }
  }

  private maxRowNumber(): number {
    let max = 0;
    for (const n of this.rows.keys()) if (n > max) max = n;
    return max;
  }
}
