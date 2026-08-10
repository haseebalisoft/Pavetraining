/** Shared parsing for admin bulk-upload commit bodies. */
export function parseBulkFieldRecord(
  raw: unknown,
): Record<string, string | null> {
  if (!raw || typeof raw !== "object") return {};
  const fields: Record<string, string | null> = {};
  for (const [key, value] of Object.entries(raw as Record<string, unknown>)) {
    if (value === null || value === undefined || value === "") {
      fields[key] = null;
    } else {
      fields[key] = String(value);
    }
  }
  return fields;
}
