import "server-only";

import * as XLSX from "xlsx";

import { ValidationError } from "@/lib/services/validationService";
import { normalizeDateValue } from "@/lib/utils/ukDate";

export type ParsedSpreadsheet = {
  headers: string[];
  rows: Array<Record<string, string | null>>;
};

export { normalizeDateValue };

function cellToString(value: unknown): string | null {
  if (value === null || value === undefined) return null;
  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) return null;
    // SheetJS date-only cells are often UTC midnight, or local midnight encoded as
    // a prior UTC evening (e.g. 1971-12-31T19:00Z for 01/01/1972 in UTC+5).
    // Prefer local calendar day when UTC time is afternoon/evening and local
    // date is ahead; otherwise use UTC Y-M-D (Vercel is UTC).
    const utcY = value.getUTCFullYear();
    const utcM = value.getUTCMonth() + 1;
    const utcD = value.getUTCDate();
    const localY = value.getFullYear();
    const localM = value.getMonth() + 1;
    const localD = value.getDate();
    const useLocal =
      value.getUTCHours() >= 12 &&
      (localY > utcY ||
        localM > utcM ||
        (localY === utcY && localM === utcM && localD > utcD));
    const year = useLocal ? localY : utcY;
    const month = String(useLocal ? localM : utcM).padStart(2, "0");
    const day = String(useLocal ? localD : utcD).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }
  if (typeof value === "number") {
    // Prefer normalizing Excel serials immediately to YYYY-MM-DD.
    if (value > 20000 && value < 60000) {
      const ms = Date.UTC(1899, 11, 30) + Math.round(value) * 86_400_000;
      const parsed = new Date(ms);
      if (!Number.isNaN(parsed.getTime())) {
        const year = parsed.getUTCFullYear();
        const month = String(parsed.getUTCMonth() + 1).padStart(2, "0");
        const day = String(parsed.getUTCDate()).padStart(2, "0");
        return `${year}-${month}-${day}`;
      }
    }
    return String(value);
  }
  const text = String(value).trim();
  if (!text) return null;
  // SharePoint / Excel blank markers used in matrix exports
  if (/^(—|–|-|n\/?a|null|none)$/i.test(text)) return null;
  return text;
}

function normalizeHeader(header: string): string {
  return header.trim().toLowerCase().replace(/\s+/g, " ");
}

/**
 * Parse .xlsx or .csv bytes into row objects keyed by original header labels.
 */
export function parseSpreadsheetBuffer(
  bytes: ArrayBuffer | Uint8Array,
  fileName: string,
): ParsedSpreadsheet {
  const lower = fileName.toLowerCase();
  if (!lower.endsWith(".xlsx") && !lower.endsWith(".csv")) {
    throw new ValidationError("Only .xlsx and .csv files are supported.");
  }

  const workbook = XLSX.read(bytes, {
    type: "array",
    // Keep Excel serial day numbers (not JS Dates) — Dates shift a calendar day
    // in non-UTC timezones. cellToString + normalizeDateValue handle serials.
    cellDates: false,
    raw: true,
  });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) {
    throw new ValidationError("Spreadsheet has no sheets.");
  }
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) {
    throw new ValidationError("Spreadsheet sheet could not be read.");
  }

  const matrix = XLSX.utils.sheet_to_json<(string | number | Date | null)[]>(
    sheet,
    {
      header: 1,
      defval: null,
      raw: true,
      blankrows: false,
    },
  );

  if (!matrix.length) {
    throw new ValidationError("Spreadsheet is empty.");
  }

  const headerRow = matrix[0] ?? [];
  const headers = headerRow
    .map((cell) => (cell === null || cell === undefined ? "" : String(cell).trim()))
    .filter((h) => h.length > 0);

  if (!headers.length) {
    throw new ValidationError("Spreadsheet has no header row.");
  }

  const rows: Array<Record<string, string | null>> = [];
  for (let i = 1; i < matrix.length; i += 1) {
    const raw = matrix[i] ?? [];
    const row: Record<string, string | null> = {};
    let hasAny = false;
    for (let c = 0; c < headers.length; c += 1) {
      const header = headers[c]!;
      const value = cellToString(raw[c]);
      row[header] = value;
      if (value) hasAny = true;
    }
    if (hasAny) rows.push(row);
  }

  if (!rows.length) {
    throw new ValidationError("Spreadsheet has no data rows.");
  }

  return { headers, rows };
}

export function pickField(
  row: Record<string, string | null>,
  aliases: string[],
): string | null {
  const map = new Map<string, string | null>();
  for (const [key, value] of Object.entries(row)) {
    map.set(normalizeHeader(key), value);
  }
  for (const alias of aliases) {
    const hit = map.get(normalizeHeader(alias));
    if (hit) return hit;
  }
  return null;
}

/**
 * Normalize a commit request's `source` object (the ORIGINAL spreadsheet cells
 * keyed by exact Excel header).
 *
 * The Training Matrix importer needs these raw display headers: `fields` only
 * carries the ~6 mapped meta columns, so without `source` every N-code expiry
 * column and every category record is silently dropped at commit.
 */
export function normalizeSourceRow(
  value: unknown,
): Record<string, string | null> | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return undefined;
  }
  const source: Record<string, string | null> = {};
  for (const [key, raw] of Object.entries(value as Record<string, unknown>)) {
    source[key] =
      raw === null || raw === undefined || raw === "" ? null : String(raw);
  }
  return source;
}
