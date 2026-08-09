/**
 * Format an event duration from start/end timestamps for customer UI.
 * Same-day → hours/minutes; multi-day → whole days (inclusive calendar span).
 */
export function formatEventDuration(
  start: string | Date | null | undefined,
  end: string | Date | null | undefined,
): string | null {
  if (!start || !end) return null;
  const startDate = start instanceof Date ? start : new Date(start);
  const endDate = end instanceof Date ? end : new Date(end);
  if (Number.isNaN(startDate.getTime()) || Number.isNaN(endDate.getTime())) {
    return null;
  }
  if (endDate.getTime() < startDate.getTime()) return null;

  const startDay = Date.UTC(
    startDate.getFullYear(),
    startDate.getMonth(),
    startDate.getDate(),
  );
  const endDay = Date.UTC(
    endDate.getFullYear(),
    endDate.getMonth(),
    endDate.getDate(),
  );
  const daySpan = Math.round((endDay - startDay) / 86_400_000) + 1;

  if (daySpan > 1) {
    return daySpan === 2 ? "2 days" : `${daySpan} days`;
  }

  const minutes = Math.round(
    (endDate.getTime() - startDate.getTime()) / 60_000,
  );
  if (minutes <= 0) return null;
  const hours = Math.floor(minutes / 60);
  const mins = minutes % 60;
  if (hours === 0) return mins === 1 ? "1 minute" : `${mins} minutes`;
  if (mins === 0) return hours === 1 ? "1 hour" : `${hours} hours`;
  return `${hours}h ${mins}m`;
}
