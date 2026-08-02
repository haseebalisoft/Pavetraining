export type DateValue = string | Date | null | undefined;

const ISO_DATE_ONLY =
  /^(\d{4})-(\d{2})-(\d{2})(?:T00:00(?::00(?:\.0+)?)?(?:Z|[+-]\d{2}:?\d{2})?)?$/;
const UK_DATE_ONLY = /^(\d{2})\/(\d{2})\/(\d{4})$/;

const UK_DATE_FORMATTER = new Intl.DateTimeFormat("en-GB", {
  day: "2-digit",
  month: "2-digit",
  year: "numeric",
});

const UK_TIME_FORMATTER = new Intl.DateTimeFormat("en-GB", {
  hour: "2-digit",
  minute: "2-digit",
  hour12: false,
});

const UK_MONTH_YEAR_FORMATTER = new Intl.DateTimeFormat("en-GB", {
  month: "long",
  year: "numeric",
});

const UK_DAY_FORMATTER = new Intl.DateTimeFormat("en-GB", {
  day: "2-digit",
});

const UK_SHORT_MONTH_FORMATTER = new Intl.DateTimeFormat("en-GB", {
  month: "short",
});

function invalidValueFallback(value: DateValue, fallback: string): string {
  return typeof value === "string" && value.trim() ? value : fallback;
}

function parseDateValue(value: DateValue): Date | null {
  if (value instanceof Date) {
    return Number.isNaN(value.getTime()) ? null : value;
  }
  if (!value?.trim()) {
    return null;
  }
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function isValidCalendarDate(year: number, month: number, day: number): boolean {
  const parsed = new Date(Date.UTC(year, month - 1, day));
  return (
    parsed.getUTCFullYear() === year &&
    parsed.getUTCMonth() === month - 1 &&
    parsed.getUTCDate() === day
  );
}

/**
 * Formats display dates in the app-wide numeric UK style: DD/MM/YYYY.
 * Exact ISO calendar dates are formatted from their components so a timezone
 * conversion cannot move a date of birth or expiry into the previous day.
 */
export function formatDate(value: DateValue, fallback = "—"): string {
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (!trimmed) return fallback;

    const ukMatch = UK_DATE_ONLY.exec(trimmed);
    if (ukMatch) {
      const [, day, month, year] = ukMatch;
      return isValidCalendarDate(Number(year), Number(month), Number(day))
        ? trimmed
        : value;
    }

    const match = ISO_DATE_ONLY.exec(trimmed);
    if (match) {
      const [, year, month, day] = match;
      const isValid = isValidCalendarDate(
        Number(year),
        Number(month),
        Number(day),
      );
      return isValid ? `${day}/${month}/${year}` : value;
    }
  }

  const parsed = parseDateValue(value);
  return parsed
    ? UK_DATE_FORMATTER.format(parsed)
    : invalidValueFallback(value, fallback);
}

/** Formats a valid date/time value as 24-hour HH:mm, otherwise returns null. */
export function formatTime(value: DateValue): string | null {
  const parsed = parseDateValue(value);
  return parsed ? UK_TIME_FORMATTER.format(parsed) : null;
}

/** Formats a date/time as DD/MM/YYYY HH:mm using the shared date formatter. */
export function formatDateTime(value: DateValue, fallback = "—"): string {
  const parsed = parseDateValue(value);
  if (!parsed) {
    return invalidValueFallback(value, fallback);
  }

  return `${formatDate(value, fallback)} ${UK_TIME_FORMATTER.format(parsed)}`;
}

/** Long month heading used by calendar navigation (for example, August 2026). */
export function formatMonthYear(value: DateValue, fallback = "—"): string {
  const parsed = parseDateValue(value);
  return parsed
    ? UK_MONTH_YEAR_FORMATTER.format(parsed)
    : invalidValueFallback(value, fallback);
}

/** Two-digit day used by compact date blocks (for example, 02). */
export function formatDay(value: DateValue, fallback = "—"): string {
  const parsed = parseDateValue(value);
  return parsed
    ? UK_DAY_FORMATTER.format(parsed)
    : invalidValueFallback(value, fallback);
}

/** Abbreviated UK month used by compact date blocks (for example, Aug). */
export function formatShortMonth(
  value: DateValue,
  fallback = "TBC",
): string {
  const parsed = parseDateValue(value);
  return parsed
    ? UK_SHORT_MONTH_FORMATTER.format(parsed)
    : invalidValueFallback(value, fallback);
}
