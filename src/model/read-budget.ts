import type { ModelReadBudget } from "../read-limits.js";

// Scoped internal guards avoid exposing import state on the public worksheet model. They are
// removed before the workbook is returned, so subsequent user edits are not subject to read limits.
const budgets = new WeakMap<object, ModelReadBudget>();

export function modelReadBudget(host: object): ModelReadBudget | undefined {
  return budgets.get(host);
}

export function withModelReadBudget<T>(host: object, budget: ModelReadBudget, read: () => T): T {
  const previous = budgets.get(host);
  budgets.set(host, budget);
  try {
    return read();
  } finally {
    if (previous) budgets.set(host, previous);
    else budgets.delete(host);
  }
}
