import "server-only";

/** Portal booking lifecycle — offered dates vs confirmed booking. */
export type BookingStatus = "Tentative" | "Confirmed";

/** Values written to SharePoint Events.FreeBusy (calendar Free/Busy). */
export type FreeBusyValue = "Tentative" | "Busy";

/** Outlook Graph showAs for the synced calendar event. */
export type OutlookShowAs = "tentative" | "busy";

export function normalizeBookingStatus(value: unknown): BookingStatus | null {
  if (value === null || value === undefined || value === "") return null;
  const text = String(value).trim().toLowerCase();
  if (!text) return null;
  if (
    text === "confirmed" ||
    text === "busy" ||
    text === "confirm" ||
    text === "booked"
  ) {
    return "Confirmed";
  }
  if (
    text === "tentative" ||
    text === "offered" ||
    text === "proposed" ||
    text === "draft"
  ) {
    return "Tentative";
  }
  return null;
}

export function bookingStatusFromFreeBusy(value: unknown): BookingStatus {
  return normalizeBookingStatus(value) ?? "Tentative";
}

export function freeBusyFromBookingStatus(
  status: BookingStatus,
): FreeBusyValue {
  return status === "Confirmed" ? "Busy" : "Tentative";
}

export function outlookShowAsFromBookingStatus(
  status: BookingStatus,
): OutlookShowAs {
  return status === "Confirmed" ? "busy" : "tentative";
}

export function shouldSendBookingConfirmation(input: {
  previous: BookingStatus | null | undefined;
  next: BookingStatus;
}): boolean {
  return input.next === "Confirmed" && input.previous !== "Confirmed";
}
