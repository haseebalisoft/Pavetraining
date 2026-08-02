export type DisplayDateValue = string | null | undefined;

interface CalendarDateParts {
  year: number;
  month: number;
  day: number;
}

const DATE_ONLY_ISO =
  /^(\d{4})-(\d{2})-(\d{2})(?:T00:00(?::00(?:\.0+)?)?(?:Z|[+-]\d{2}:?\d{2})?)?$/;
const UK_DATE_ONLY = /^(\d{2})\/(\d{2})\/(\d{4})$/;

function pad2(value: number): string {
  return value < 10 ? "0" + value : String(value);
}

function isValidCalendarDate(parts: CalendarDateParts): boolean {
  if (parts.month < 1 || parts.month > 12 || parts.day < 1 || parts.day > 31) {
    return false;
  }

  const probe = new Date(Date.UTC(parts.year, parts.month - 1, parts.day));
  return (
    probe.getUTCFullYear() === parts.year &&
    probe.getUTCMonth() === parts.month - 1 &&
    probe.getUTCDate() === parts.day
  );
}

function dateOnlyParts(value: string): CalendarDateParts | null {
  const match = DATE_ONLY_ISO.exec(value);
  if (!match) return null;

  const parts: CalendarDateParts = {
    year: Number(match[1]),
    month: Number(match[2]),
    day: Number(match[3]),
  };
  return isValidCalendarDate(parts) ? parts : null;
}

function localDateParts(value: Date): CalendarDateParts {
  return {
    year: value.getFullYear(),
    month: value.getMonth() + 1,
    day: value.getDate(),
  };
}

function formatDateParts(parts: CalendarDateParts): string {
  return pad2(parts.day) + "/" + pad2(parts.month) + "/" + parts.year;
}

function parseDate(value: string): Date | null {
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? null : parsed;
}

/** Display a date as DD/MM/YYYY without shifting date-only ISO values. */
export function formatDate(value: DisplayDateValue): string {
  const raw = typeof value === "string" ? value.trim() : "";
  if (!raw) return "";

  const ukMatch = UK_DATE_ONLY.exec(raw);
  if (ukMatch) {
    const ukParts: CalendarDateParts = {
      year: Number(ukMatch[3]),
      month: Number(ukMatch[2]),
      day: Number(ukMatch[1]),
    };
    return isValidCalendarDate(ukParts) ? raw : value || "";
  }

  const calendarParts = dateOnlyParts(raw);
  if (calendarParts) return formatDateParts(calendarParts);

  const parsed = parseDate(raw);
  return parsed ? formatDateParts(localDateParts(parsed)) : raw;
}

/** Display a time as HH:mm in the browser's local timezone. */
export function formatTime(value: DisplayDateValue): string {
  const raw = typeof value === "string" ? value.trim() : "";
  if (!raw) return "";

  if (dateOnlyParts(raw)) return "00:00";

  const parsed = parseDate(raw);
  return parsed ? pad2(parsed.getHours()) + ":" + pad2(parsed.getMinutes()) : raw;
}

/** Display a timestamp as DD/MM/YYYY HH:mm in the browser's local timezone. */
export function formatDateTime(value: DisplayDateValue): string {
  const raw = typeof value === "string" ? value.trim() : "";
  if (!raw) return "";

  const calendarParts = dateOnlyParts(raw);
  if (calendarParts) return formatDateParts(calendarParts) + " 00:00";

  const parsed = parseDate(raw);
  return parsed
    ? formatDateParts(localDateParts(parsed)) +
        " " +
        pad2(parsed.getHours()) +
        ":" +
        pad2(parsed.getMinutes())
    : raw;
}
