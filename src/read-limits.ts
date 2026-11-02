import { ReadLimitError } from "./errors.js";

/** ZIP allocation budgets, in bytes except maxEntries. Infinity explicitly disables a limit. */
export interface ReadLimits {
  maxInputBytes?: number;
  maxEntries?: number;
  maxEntryBytes?: number;
  maxTotalBytes?: number;
}

export const DEFAULT_READ_LIMITS: Readonly<Required<ReadLimits>> = Object.freeze({
  maxInputBytes: 128 * 1024 * 1024,
  maxEntries: 10_000,
  maxEntryBytes: 256 * 1024 * 1024,
  maxTotalBytes: 1024 * 1024 * 1024,
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
