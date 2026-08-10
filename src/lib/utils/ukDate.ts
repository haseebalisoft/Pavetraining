/**
 * UK-first date parsing for bulk upload + forms.
 * Pure helper (no server-only) so unit scripts can import it.
 */

/** Excel serial day number → YYYY-MM-DD (Excel epoch 1899-12-30). */
function excelSerialToIso(serial: number): string | null {
  if (!Number.isFinite(serial) || serial <= 20000 || serial >= 60000) {
    return null;
  }
  const ms = Date.UTC(1899, 11, 30) + Math.round(serial) * 86_400_000;
  const parsed = new Date(ms);
  if (Number.isNaN(parsed.getTime())) return null;
  const year = parsed.getUTCFullYear();
  const month = String(parsed.getUTCMonth() + 1).padStart(2, "0");
  const day = String(parsed.getUTCDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

/** Normalize common date cell values to YYYY-MM-DD when possible (UK-first). */
export function normalizeDateValue(value: string | null): string | null {
  if (!value?.trim()) return null;
  const text = value.trim();
  if (/^(—|–|-|n\/?a|null|none)$/i.test(text)) return null;

  if (/^\d{4}-\d{2}-\d{2}/.test(text)) {
    return text.slice(0, 10);
  }

  const uk = text.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
  if (uk) {
    const day = Number(uk[1]);
    const month = Number(uk[2]);
    const year = Number(uk[3]);
    if (month < 1 || month > 12 || day < 1 || day > 31) {
      return null;
    }
    const probe = new Date(Date.UTC(year, month - 1, day));
    if (
      probe.getUTCFullYear() !== year ||
      probe.getUTCMonth() !== month - 1 ||
      probe.getUTCDate() !== day
    ) {
      return null;
    }
    return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
  }

  const short = text.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2})$/);
  if (short) {
    const first = Number(short[1]);
    const second = Number(short[2]);
    let year = Number(short[3]);
    year += year >= 70 ? 1900 : 2000;
    let day: number;
    let month: number;
    if (second > 12 && first <= 12) {
      month = first;
      day = second;
    } else {
      day = first;
      month = second;
    }
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      const probe = new Date(Date.UTC(year, month - 1, day));
      if (
        probe.getUTCFullYear() === year &&
        probe.getUTCMonth() === month - 1 &&
        probe.getUTCDate() === day
      ) {
        return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
      }
    }
    return null;
  }

  if (/^\d+(\.\d+)?$/.test(text)) {
    return excelSerialToIso(Number(text));
  }

  if (/[a-zA-Z]/.test(text)) {
    const ms = Date.parse(text);
    if (!Number.isNaN(ms)) {
      const parsed = new Date(ms);
      const year = parsed.getFullYear();
      const month = String(parsed.getMonth() + 1).padStart(2, "0");
      const day = String(parsed.getDate()).padStart(2, "0");
      return `${year}-${month}-${day}`;
    }
  }

  return null;
}
