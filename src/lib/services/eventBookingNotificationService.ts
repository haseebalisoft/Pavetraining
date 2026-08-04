import "server-only";

import {
  shouldSendBookingConfirmation,
  type BookingStatus,
} from "@/lib/services/bookingStatusService";
import { buildBookingIcsInvite } from "@/lib/services/calendar/icsInviteService";
import { getNotificationSettings } from "@/lib/services/notificationConfig";
import { hasRecentNotificationDedupe } from "@/lib/services/notificationLogService";
import { resolveTrainingManagerRecipients } from "@/lib/services/notificationRecipientService";
import { sendNotification } from "@/lib/services/notificationService";
import { bookingConfirmedEmailTemplate } from "@/lib/services/notificationTemplateService";
import { formatDate, formatTime } from "@/lib/utils/formatDate";
import type { NotificationSendResult } from "@/types/notifications";

export type BookingConfirmationResult = {
  attempted: boolean;
  skipped: boolean;
  skipReason?: string;
  recipients: string[];
  results: NotificationSendResult[];
};

function timeRangeLabel(
  start: string | null | undefined,
  end: string | null | undefined,
): string {
  const startTime = formatTime(start);
  const endTime = formatTime(end);
  if (startTime && endTime) return `${startTime} – ${endTime}`;
  return startTime ?? endTime ?? "Time TBC";
}

/**
 * Emails company Training Managers when a booking becomes Confirmed.
 * Attaches an Outlook-compatible .ics invite so Accept adds it to calendar.
 * Never throws — failures are logged via sendNotification.
 */
export async function notifyBookingConfirmed(input: {
  eventId: string;
  title: string;
  companyId: string | null | undefined;
  companyName?: string | null;
  eventDate?: string | null;
  endDate?: string | null;
  location?: string | null;
  previousStatus: BookingStatus | null | undefined;
  nextStatus: BookingStatus;
  actorEmail?: string | null;
}): Promise<BookingConfirmationResult> {
  if (
    !shouldSendBookingConfirmation({
      previous: input.previousStatus,
      next: input.nextStatus,
    })
  ) {
    return {
      attempted: false,
      skipped: true,
      skipReason:
        "Booking is not newly confirmed. Set status to Confirmed to email Training Managers.",
      recipients: [],
      results: [],
    };
  }

  if (!input.companyId?.trim()) {
    return {
      attempted: false,
      skipped: true,
      skipReason: "Company is required to notify Training Managers.",
      recipients: [],
      results: [],
    };
  }

  const managers = await resolveTrainingManagerRecipients(input.companyId);
  if (managers.length === 0) {
    return {
      attempted: true,
      skipped: true,
      skipReason:
        "No Active Training Managers found for this company (with customer notifications enabled).",
      recipients: [],
      results: [],
    };
  }

  const dateLabel = formatDate(input.eventDate) || "Date TBC";
  const timeLabel = timeRangeLabel(input.eventDate, input.endDate);
  const settings = await getNotificationSettings();
  const organizerEmail =
    settings.fromEmail?.trim() || "info@pavetraining.co.uk";
  const results: NotificationSendResult[] = [];
  const recipients: string[] = [];

  for (const manager of managers) {
    const dedupeKey = `booking-confirmed:${input.eventId}:${manager.email}`;
    if (await hasRecentNotificationDedupe(dedupeKey, 2)) {
      continue;
    }

    const template = bookingConfirmedEmailTemplate({
      trainingManagerName: manager.displayName,
      eventTitle: input.title,
      dateLabel,
      timeLabel,
      location: input.location,
    });

    const ics = buildBookingIcsInvite({
      eventId: input.eventId,
      title: input.title,
      startIso: input.eventDate,
      endIso: input.endDate,
      location: input.location,
      description: [
        `Event: ${input.title}`,
        `Date: ${dateLabel}`,
        `Time: ${timeLabel}`,
        input.location?.trim() ? `Location: ${input.location.trim()}` : null,
        input.companyName?.trim()
          ? `Company: ${input.companyName.trim()}`
          : null,
      ]
        .filter(Boolean)
        .join("\n"),
      organizerEmail,
      attendeeEmail: manager.email,
      attendeeName: manager.displayName,
    });

    recipients.push(manager.email);
    results.push(
      await sendNotification({
        type: "booking_confirmed",
        to: manager.email,
        subject: template.subject,
        text: template.text,
        html: template.html,
        companyName: input.companyName ?? manager.companyName,
        itemId: input.eventId,
        dedupeKey,
        actorEmail: input.actorEmail,
        detail: `Booking confirmed: ${input.title}`,
        fromName: "PAVE Training",
        attachments: ics
          ? [
              {
                filename: ics.filename,
                contentType: ics.contentType,
                content: ics.content,
                encoding: "utf8",
              },
            ]
          : undefined,
      }),
    );
  }

  const sentCount = results.filter((r) => r.status === "sent").length;
  const failed = results.find(
    (r) => r.status === "failed" || r.status === "not_configured",
  );

  return {
    attempted: true,
    skipped: recipients.length === 0 || sentCount === 0,
    skipReason:
      recipients.length === 0
        ? "Confirmation already sent recently for these recipients."
        : sentCount === 0
          ? failed?.errorMessage || "Booking confirmation email was not delivered."
          : undefined,
    recipients,
    results,
  };
}
