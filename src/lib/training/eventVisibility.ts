/**
 * The live Events list has TWO customer-visibility columns:
 *   "Customer Visible" (Customer_x0020_Visible) — the one the app writes
 *   "CustomerVisible"  (CustomerVisible)        — a later duplicate
 * Customers used to see an event when EITHER was true, so unticking
 * "Customer visible" in the app (which wrote only the first) left the event
 * visible through the second. Now both are written together, and an event is
 * visible only when every column that is present says so — when they
 * disagree, hiding wins.
 */

export const EVENT_VISIBILITY_COLUMNS = [
  "Customer_x0020_Visible",
  "CustomerVisible",
] as const;

function truthy(value: unknown): boolean {
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value !== 0;
  const text = String(value ?? "").trim().toLowerCase();
  return ["true", "yes", "1"].includes(text);
}

export function isEventCustomerVisible(
  fields: Record<string, unknown>,
): boolean {
  const present = EVENT_VISIBILITY_COLUMNS.map((name) => fields[name]).filter(
    (value) => value !== undefined && value !== null && value !== "",
  );
  return present.length > 0 && present.every(truthy);
}

/** Both visibility columns, for a create/update payload. */
export function eventVisibilityFields(
  visible: boolean,
): Record<(typeof EVENT_VISIBILITY_COLUMNS)[number], boolean> {
  return { Customer_x0020_Visible: visible, CustomerVisible: visible };
}
