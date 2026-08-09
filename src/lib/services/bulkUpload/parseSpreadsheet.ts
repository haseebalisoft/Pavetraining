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
    // Use local Y-M-D — toISOString() shifts calendar day near midnight in US timezones.
    const year = value.getFullYear();
    const month = String(value.getMonth() + 1).padStart(2, "0");
    const day = String(value.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }
  if (typeof value === "number") {
    // Excel serial dates often arrive as numbers; leave as string of number
    // unless it looks like a date serial — callers normalize DOB separately.
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
    cellDates: true,
    raw: false,
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
      raw: false,
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
