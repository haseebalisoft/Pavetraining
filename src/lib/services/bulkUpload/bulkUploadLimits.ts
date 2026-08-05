/**
 * Hard row limits for bulk upload, by import type.
 *
 * Pure + dependency-free (only a type-erased import) so it can be shared by the
 * server (preview/commit) AND the admin UI (client-side mirror), and exercised
 * by scripts/test-bulk-upload-limits.mjs under Node without Graph.
 *
 * Wayne's requirement: workforce imports accept at most 50 records per upload.
 * Only populated data rows count — the header row and blank rows are excluded.
 */
import type { BulkImportType } from "@/types/bulkUpload";

/**
 * Max populated data rows accepted per upload, by import type. A type with no
 * entry has no row limit. Company has no limit today; give it its own here
 * (e.g. `company: 200`) if a separate cap is ever required — callers and the UI
 * pick it up automatically.
 */
export const BULK_UPLOAD_ROW_LIMITS: Partial<Record<BulkImportType, number>> = {
  workforce: 50,
};

/** Human labels used when building the limit error message. */
const IMPORT_TYPE_LABELS: Partial<Record<BulkImportType, string>> = {
  company: "Company",
  workforce: "Workforce",
  trainingMatrix: "Training Matrix",
  npors: "NPORS",
  eusr: "EUSR",
  streetworks: "Streetworks",
  inHouse: "In-House",
  nvq: "NVQ",
};

/** The configured row limit for a type, or null when there is no limit. */
export function bulkUploadRowLimit(importType: BulkImportType): number | null {
  return BULK_UPLOAD_ROW_LIMITS[importType] ?? null;
}

/** Clear, user-facing message when an upload exceeds its row limit. */
export function bulkUploadRowLimitError(
  importType: BulkImportType,
  limit: number,
): string {
  const label = IMPORT_TYPE_LABELS[importType] ?? "This";
  return `${label} bulk upload supports a maximum of ${limit} records per upload.`;
}

export type BulkUploadRowLimitCheck = {
  ok: boolean;
  limit: number | null;
  error: string | null;
};

/**
 * Enforce the row limit for a type against a count of POPULATED data rows.
 * Pass a count that already excludes the header and blank rows
 * (see {@link countPopulatedRows}); the parser strips those before we get here.
 */
export function checkBulkUploadRowLimit(
  importType: BulkImportType,
  dataRowCount: number,
): BulkUploadRowLimitCheck {
  const limit = bulkUploadRowLimit(importType);
  if (limit !== null && dataRowCount > limit) {
    return {
      ok: false,
      limit,
      error: bulkUploadRowLimitError(importType, limit),
    };
  }
  return { ok: true, limit, error: null };
}

/** True when every field in the set is blank/absent (an "empty" row). */
export function isEmptyFieldSet(
  fields: Record<string, string | null> | null | undefined,
): boolean {
  if (!fields) return true;
  return !Object.values(fields).some(
    (value) => value !== null && value !== undefined && String(value).trim() !== "",
  );
}

/**
 * Count only populated rows — blank rows (and, by construction, the header,
 * which is never a data row) do not count toward the limit.
 */
export function countPopulatedRows(
  rows: Array<{ fields?: Record<string, string | null> | null }>,
): number {
  return rows.reduce(
    (total, row) => (isEmptyFieldSet(row.fields) ? total : total + 1),
    0,
  );
}
