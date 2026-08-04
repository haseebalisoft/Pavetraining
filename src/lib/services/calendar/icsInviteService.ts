import "server-only";

/**
 * Build a minimal METHOD:REQUEST .ics so mail clients (Outlook) can
 * Accept → add the booking to the recipient's calendar.
 */
export function buildBookingIcsInvite(input: {
  eventId: string;
  title: string;
  startIso: string | null | undefined;
  endIso: string | null | undefined;
  location?: string | null;
  description?: string | null;
  organizerEmail: string;
  attendeeEmail: string;
  attendeeName?: string | null;
}): { filename: string; contentType: string; content: string } | null {
  const start = parseUtc(input.startIso);
  if (!start) return null;
  const end =
    parseUtc(input.endIso) ??
    new Date(start.getTime() + 60 * 60 * 1000);

  const uid = `pave-booking-${sanitizeUid(input.eventId)}@pavetraining.co.uk`;
  const stamp = formatIcsUtc(new Date());
  const title = foldIcsText(escapeIcsText(input.title.trim() || "PAVE Training"));
  const location = input.location?.trim()
    ? foldIcsText(escapeIcsText(input.location.trim()))
    : null;
  const description = input.description?.trim()
    ? foldIcsText(escapeIcsText(input.description.trim()))
    : foldIcsText(
        escapeIcsText("Training booking confirmed via PAVE Training Portal."),
      );
  const organizer = input.organizerEmail.trim().toLowerCase();
  const attendee = input.attendeeEmail.trim().toLowerCase();
  const attendeeCn = input.attendeeName?.trim()
    ? `;CN=${escapeIcsText(input.attendeeName.trim())}`
    : "";

  const lines = [
    "BEGIN:VCALENDAR",
    "PRODID:-//PAVE Training//Portal//EN",
    "VERSION:2.0",
    "CALSCALE:GREGORIAN",
    "METHOD:REQUEST",
    "BEGIN:VEVENT",
    `UID:${uid}`,
    `DTSTAMP:${stamp}`,
    `DTSTART:${formatIcsUtc(start)}`,
    `DTEND:${formatIcsUtc(end)}`,
    `SUMMARY:${title}`,
    `DESCRIPTION:${description}`,
    location ? `LOCATION:${location}` : null,
    `ORGANIZER;CN=PAVE Training:mailto:${organizer}`,
    `ATTENDEE${attendeeCn};ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE:mailto:${attendee}`,
    "STATUS:CONFIRMED",
    "SEQUENCE:0",
    "TRANSP:OPAQUE",
    "END:VEVENT",
    "END:VCALENDAR",
  ].filter((line): line is string => Boolean(line));

  return {
    filename: "pave-training-booking.ics",
    contentType: "text/calendar; method=REQUEST; charset=utf-8",
    content: `${lines.join("\r\n")}\r\n`,
  };
}

function parseUtc(value: string | null | undefined): Date | null {
  if (!value?.trim()) return null;
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return null;
  return date;
}

function formatIcsUtc(date: Date): string {
  const y = date.getUTCFullYear();
  const m = String(date.getUTCMonth() + 1).padStart(2, "0");
  const d = String(date.getUTCDate()).padStart(2, "0");
  const hh = String(date.getUTCHours()).padStart(2, "0");
  const mm = String(date.getUTCMinutes()).padStart(2, "0");
  const ss = String(date.getUTCSeconds()).padStart(2, "0");
  return `${y}${m}${d}T${hh}${mm}${ss}Z`;
}

function escapeIcsText(value: string): string {
  return value
    .replace(/\\/g, "\\\\")
    .replace(/;/g, "\\;")
    .replace(/,/g, "\\,")
    .replace(/\r\n|\n|\r/g, "\\n");
}

/** RFC 5545 line folding at 75 octets (approx chars for ASCII payloads). */
function foldIcsText(value: string): string {
  if (value.length <= 75) return value;
  const parts: string[] = [];
  let remaining = value;
  let first = true;
  while (remaining.length > 0) {
    const limit = first ? 75 : 74;
    parts.push(remaining.slice(0, limit));
    remaining = remaining.slice(limit);
    first = false;
  }
  return parts.join("\r\n ");
}

function sanitizeUid(value: string): string {
  return value.replace(/[^a-zA-Z0-9_-]/g, "").slice(0, 64) || "event";
}
