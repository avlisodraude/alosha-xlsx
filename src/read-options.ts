import { ImportWarningError } from "./errors.js";
import type { ReadLimits } from "./read-limits.js";

export interface ImportWarning {
  readonly code:
    | "UNSUPPORTED_FEATURE"
    | "MISSING_PART"
    | "UNRESOLVED_RELATIONSHIP"
    | "INVALID_REFERENCE"
    | "MALFORMED_PART";
  readonly part: string;
  readonly feature: string;
  readonly message: string;
  readonly ref?: string;
}

export interface ReadOptions extends ReadLimits {
  /** Report known content loss without changing the default tolerant import behavior. */
  onWarning?: (warning: ImportWarning) => void;
  /** Throw ImportWarningError after reporting the first warning. Not a full schema validator. */
  strict?: boolean;
}

export type WarningReporter = (warning: ImportWarning) => void;

export function warningReporter(options?: ReadOptions): WarningReporter | undefined {
  if (options?.strict !== undefined && typeof options.strict !== "boolean")
    throw new TypeError("strict must be a boolean");
  if (options?.onWarning !== undefined && typeof options.onWarning !== "function")
    throw new TypeError("onWarning must be a function");
  if (!options?.strict && !options?.onWarning) return undefined;
  const strict = options.strict === true;
  const onWarning = options.onWarning;
  const seen = new Set<string>();
  return (value) => {
    const key = JSON.stringify(value);
    if (seen.has(key)) return;
    seen.add(key);
    const warning = Object.freeze({ ...value });
    onWarning?.(warning);
    if (strict) throw new ImportWarningError(warning);
  };
}
