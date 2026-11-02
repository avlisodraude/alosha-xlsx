import { ReadLimitError } from "./errors.js";

/** Workbook-wide import budgets. Byte limits apply to ZIP data; model limits count allocations/work.
 * Infinity explicitly disables a limit. */
export interface ReadLimits {
  maxInputBytes?: number;
  maxEntries?: number;
  maxEntryBytes?: number;
  maxTotalBytes?: number;
  maxCells?: number;
  maxRows?: number;
  maxMergeCells?: number;
  maxColumnDefinitions?: number;
}

export const DEFAULT_READ_LIMITS: Readonly<Required<ReadLimits>> = Object.freeze({
  maxInputBytes: 128 * 1024 * 1024,
  maxEntries: 10_000,
  maxEntryBytes: 256 * 1024 * 1024,
  maxTotalBytes: 1024 * 1024 * 1024,
  maxCells: 1_000_000,
  maxRows: 1_048_576,
  maxMergeCells: 100_000,
  maxColumnDefinitions: 100_000,
});

export function resolveReadLimits(options?: ReadLimits): Required<ReadLimits> {
  const limits = { ...DEFAULT_READ_LIMITS };
  for (const key of Object.keys(limits) as Array<keyof ReadLimits>) {
    const value = options?.[key];
    if (value === undefined) continue;
    if (value !== Number.POSITIVE_INFINITY && (!Number.isSafeInteger(value) || value < 0)) {
      throw new RangeError(`${key} must be a nonnegative safe integer or Infinity`);
    }
    limits[key] = value;
  }
  return limits;
}

export function checkReadLimit(limit: keyof ReadLimits, maximum: number, actual: number): void {
  if (actual > maximum) throw new ReadLimitError(limit, maximum, actual);
}

type ModelLimit = "maxCells" | "maxRows" | "maxMergeCells" | "maxColumnDefinitions";

/** Internal counter shared by all worksheets in a single import, never by separate reads. */
export class ModelReadBudget {
  private readonly used: Record<ModelLimit, number> = {
    maxCells: 0,
    maxRows: 0,
    maxMergeCells: 0,
    maxColumnDefinitions: 0,
  };

  constructor(private readonly limits: Required<ReadLimits>) {}

  claim(limit: ModelLimit, count = 1): void {
    const actual = this.used[limit] + count;
    checkReadLimit(limit, this.limits[limit], actual);
    this.used[limit] = actual;
  }
}
