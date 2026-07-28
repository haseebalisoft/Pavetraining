import "server-only";

import * as XLSX from "xlsx";

import { ValidationError } from "@/lib/services/validationService";

export type ParsedSpreadsheet = {
  headers: string[];
  rows: Array<Record<string, string | null>>;
};

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

/** Normalize common date cell values to YYYY-MM-DD when possible. */
export function normalizeDateValue(value: string | null): string | null {
  if (!value?.trim()) return null;
  const text = value.trim();
  if (/^(—|–|-|n\/?a|null|none)$/i.test(text)) return null;

  // Already ISO-ish
  if (/^\d{4}-\d{2}-\d{2}/.test(text)) {
    return text.slice(0, 10);
  }

  // DD/MM/YYYY or DD-MM-YYYY (4-digit year — treat as UK day/month)
  const uk = text.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
  if (uk) {
    const day = uk[1]!.padStart(2, "0");
    const month = uk[2]!.padStart(2, "0");
    const year = uk[3]!;
    return `${year}-${month}-${day}`;
  }

  // M/D/YY or D/M/YY (Excel display strings from Workforce list.xlsx)
  const short = text.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2})$/);
  if (short) {
    const first = Number(short[1]);
    const second = Number(short[2]);
    let year = Number(short[3]);
    year += year >= 70 ? 1900 : 2000;
    let month: number;
    let day: number;
    if (first > 12 && second <= 12) {
      day = first;
      month = second;
    } else {
      // Default to US M/D/YY (matches Excel serials in the client template).
      month = first;
      day = second;
    }
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
    }
  }

  // Excel serial as string number
  if (/^\d+(\.\d+)?$/.test(text)) {
    const serial = Number(text);
    if (serial > 20000 && serial < 60000) {
      const parsed = XLSX.SSF.parse_date_code(serial);
      if (parsed) {
        const month = String(parsed.m).padStart(2, "0");
        const day = String(parsed.d).padStart(2, "0");
        return `${parsed.y}-${month}-${day}`;
      }
    }
  }

  const ms = Date.parse(text);
  if (!Number.isNaN(ms)) {
    const parsed = new Date(ms);
    const year = parsed.getFullYear();
    const month = String(parsed.getMonth() + 1).padStart(2, "0");
    const day = String(parsed.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }

  return text;
}
