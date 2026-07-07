// Conditional formatting: per-worksheet `<conditionalFormatting>` blocks (docs/CONDITIONAL-FORMATTING-SPEC.md §1).
//
// A conditional format is inherently range-based, so it is stored per worksheet keyed by its `ref`
// (a single address, a range, or space-separated ranges) rather than on individual cells. The shape
// mirrors ExcelJS's `worksheet.addConditionalFormatting({ ref, rules })` so the compat facade can pass
// straight through. Rules are a discriminated union keyed by `type`: the dxf-bearing rules carry an
// optional differential `style`; the self-visualising rules (colorScale/dataBar/iconSet) carry no
// style, only cfvo/color/icon fields.

import type { Style } from "./style.js";

/** A `<cfvo>` (conditional-format value object): the threshold anchoring a colorScale/dataBar/iconSet
 *  stop. `value` maps to the `val` attribute (omitted for `min`/`max`). */
export type CfValueObject = {
  type: "num" | "percent" | "max" | "min" | "percentile" | "formula";
  value?: number | string;
};

/** An ARGB colour string, e.g. `"FFF8696B"`. */
export type CfColor = string;

/** The comparison operators a `cellIs` (and the text family) rule may use. */
export type CfOperator =
  | "lessThan"
  | "lessThanOrEqual"
  | "equal"
  | "notEqual"
  | "greaterThanOrEqual"
  | "greaterThan"
  | "between"
  | "notBetween"
  | "containsText"
  | "notContains"
  | "beginsWith"
  | "endsWith";

/** Fields shared by every rule. `priority` orders rules within a worksheet (writer assigns one when
 *  unset); `stopIfTrue` is a boolean pass-through. */
export interface CfRuleBase {
  priority?: number;
  stopIfTrue?: boolean;
}

export interface CfCellIs extends CfRuleBase {
  type: "cellIs";
  operator: CfOperator;
  formulae: (string | number)[];
  style?: Partial<Style>;
}

export interface CfExpression extends CfRuleBase {
  type: "expression";
  formulae: [string];
  style?: Partial<Style>;
}

export interface CfText extends CfRuleBase {
  type: "containsText" | "notContainsText" | "beginsWith" | "endsWith";
  text: string;
  style?: Partial<Style>;
}

export interface CfTop10 extends CfRuleBase {
  type: "top10";
  rank: number;
  percent?: boolean;
  bottom?: boolean;
  style?: Partial<Style>;
}

export interface CfAverage extends CfRuleBase {
  type: "aboveAverage" | "belowAverage";
  style?: Partial<Style>;
}

export interface CfDuplicate extends CfRuleBase {
  type: "duplicateValues" | "uniqueValues";
  style?: Partial<Style>;
}

export interface CfColorScale extends CfRuleBase {
  type: "colorScale";
  cfvo: CfValueObject[];
  color: CfColor[];
}

export interface CfDataBar extends CfRuleBase {
  type: "dataBar";
  cfvo: [CfValueObject, CfValueObject];
  color: CfColor;
}

export interface CfIconSet extends CfRuleBase {
  type: "iconSet";
  iconSet: string;
  cfvo: CfValueObject[];
}

/** A single `<cfRule>`: the discriminated union keyed by `type`. */
export type ConditionalRule =
  | CfCellIs
  | CfExpression
  | CfText
  | CfTop10
  | CfAverage
  | CfDuplicate
  | CfColorScale
  | CfDataBar
  | CfIconSet;

/** One `<conditionalFormatting sqref="…">` block: a `ref` plus its ordered rules. Mirrors the object
 *  ExcelJS's `worksheet.addConditionalFormatting(...)` accepts and returns. */
export interface ConditionalFormatting {
  ref: string;
  rules: ConditionalRule[];
}
