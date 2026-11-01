/**
 * Event start/end times in 15-minute steps (09:00, 09:15, 09:30, 09:45 …).
 *
 * Values stay in the existing `YYYY-MM-DDTHH:mm` (datetime-local) form, so
 * saving and SharePoint storage are unchanged. A stored time that is not on a
 * quarter hour (e.g. 09:10 from Outlook or older data) is kept as-is and
 * offered alongside the 15-minute choices — opening and saving an event
 * never moves it.
 */

/** "00:00", "00:15", … "23:45". */
export const QUARTER_HOUR_TIMES: readonly string[] = Array.from(
  { length: 96 },
  (_, index) => {
    const hours = Math.floor(index / 4);
    const minutes = (index % 4) * 15;
    return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}`;
  },
);

/** Time choices for one input: the quarter hours, plus the stored off-grid time. */
export function quarterHourOptions(current: string | null | undefined): string[] {
  const time = (current ?? "").trim();
  if (!/^\d{2}:\d{2}$/.test(time) || QUARTER_HOUR_TIMES.includes(time)) {
    return [...QUARTER_HOUR_TIMES];
  }
  return [...QUARTER_HOUR_TIMES, time].sort();
}

/** Split `YYYY-MM-DDTHH:mm` into its date and time parts. */
export function splitDateTimeLocal(value: string | null | undefined): {
  date: string;
  time: string;
} {
  const text = (value ?? "").trim();
  const match = text.match(/^(\d{4}-\d{2}-\d{2})(?:T(\d{2}:\d{2}))?/);
  return { date: match?.[1] ?? "", time: match?.[2] ?? "" };
}

/** Join date + time back into `YYYY-MM-DDTHH:mm` ("" until a date is set). */
export function joinDateTimeLocal(date: string, time: string): string {
  if (!date) return "";
  return `${date}T${time || "09:00"}`;
}

/** The next quarter hour at or after `date` (seconds dropped). */
export function roundUpToQuarterHour(date: Date): Date {
  const next = new Date(date);
  next.setSeconds(0, 0);
  const remainder = next.getMinutes() % 15;
  if (remainder) next.setMinutes(next.getMinutes() + (15 - remainder));
  return next;
}

/**
 * Error text when an end is not strictly after its start, else null. Either
 * side blank = nothing to compare (required-ness is checked separately).
 */
export function endAfterStartError(
  start: string | null | undefined,
  end: string | null | undefined,
): string | null {
  if (!start?.trim() || !end?.trim()) return null;
  const startMs = new Date(start).getTime();
  const endMs = new Date(end).getTime();
  if (Number.isNaN(startMs) || Number.isNaN(endMs)) return null;
  return endMs > startMs ? null : "End time must be after the start time.";
}
