import { CLIENT_MATRIX_DISPLAY_HEADERS } from "@/lib/services/bulkUpload/clientTemplateHeaders";

/** Form field prefix for per-category training dates on the matrix editor. */
export const MATRIX_TRAINING_DATE_PREFIX = "training:";

/** SharePoint text column that stores `{ [header]: "YYYY-MM-DD" }` JSON. */
export const MATRIX_TRAINING_DATES_FIELD = "TrainingDates";

/** Template headers that have a training date + expiry date in the editor. */
export const MATRIX_EDITOR_DATE_HEADERS: string[] =
  CLIENT_MATRIX_DISPLAY_HEADERS.filter(
    (header) => header !== "Name" && header !== "DOB",
  );

export function matrixTrainingDateFieldName(header: string): string {
  return `${MATRIX_TRAINING_DATE_PREFIX}${header}`;
}

export function matrixTrainingDateColumnKey(header: string): string {
  return `${header} Training Date`;
}

/** ISO date stored under `header::expiry` when SharePoint has no date column. */
export const MATRIX_STORED_EXPIRY_SUFFIX = "::expiry";

export function matrixStoredExpiryKey(header: string): string {
  return `${header}${MATRIX_STORED_EXPIRY_SUFFIX}`;
}

export function isMatrixTrainingDateFieldName(name: string): boolean {
  return name.startsWith(MATRIX_TRAINING_DATE_PREFIX);
}

export function headerFromTrainingDateFieldName(name: string): string | null {
  if (!isMatrixTrainingDateFieldName(name)) return null;
  return name.slice(MATRIX_TRAINING_DATE_PREFIX.length);
}

/**
 * Expiry cells only — training-date keys must not feed "next expiry" or
 * colour filters.
 */
export function isMatrixExpiryColumnKey(key: string): boolean {
  const header = key.trim();
  if (!header || header === "Name" || header === "DOB") return false;
  if (header.startsWith(MATRIX_TRAINING_DATE_PREFIX)) return false;
  if (header.endsWith(" Training Date")) return false;
  if (header.endsWith(MATRIX_STORED_EXPIRY_SUFFIX)) return false;
  return true;
}

/** Drawer section heading for a matrix date pair. */
export function matrixEditorSectionTitle(header: string): string {
  if (header === "Face ift") return "Face Fit";
  if (header.endsWith(" Expiry")) return header.slice(0, -" Expiry".length);
  return header;
}

export function parseTrainingDatesJson(
  raw: unknown,
): Record<string, string> {
  if (raw == null || raw === "") return {};
  const text = String(raw).trim();
  if (!text) return {};
  try {
    const parsed = JSON.parse(text) as unknown;
    if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
      return {};
    }
    const out: Record<string, string> = {};
    for (const [key, value] of Object.entries(
      parsed as Record<string, unknown>,
    )) {
      if (!key.trim() || value == null || value === "") continue;
      const iso = String(value).trim().slice(0, 10);
      if (/^\d{4}-\d{2}-\d{2}$/.test(iso)) out[key] = iso;
    }
    return out;
  } catch {
    return {};
  }
}

export function serializeTrainingDatesJson(
  dates: Record<string, string | null | undefined>,
): string {
  const compact: Record<string, string> = {};
  for (const [header, value] of Object.entries(dates)) {
    const iso = value?.trim().slice(0, 10);
    if (!header.trim() || !iso || !/^\d{4}-\d{2}-\d{2}$/.test(iso)) continue;
    compact[header] = iso;
  }
  return JSON.stringify(compact);
}
