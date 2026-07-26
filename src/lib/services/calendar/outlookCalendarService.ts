import "server-only";

import { getGraphClient } from "@/lib/graph/graphClient";
import type {
  OutlookEventPayload,
  OutlookEventResult,
} from "@/types/calendarSync";

export type OutlookCalendarConfig = {
  userId: string;
  calendarId: string | null;
};

/**
 * Outlook mailbox/calendar configuration for app-only Graph access.
 * Requires OUTLOOK_USER_ID (UPN or Graph user id).
 * Optional OUTLOOK_CALENDAR_ID targets a specific calendar; otherwise default.
 */
export function getOutlookCalendarConfig(): OutlookCalendarConfig | null {
  const userId = process.env.OUTLOOK_USER_ID?.trim() || "";
  if (!userId) return null;
  const calendarId = process.env.OUTLOOK_CALENDAR_ID?.trim() || null;
  return { userId, calendarId };
}

export function isOutlookCalendarConfigured(): boolean {
  return Boolean(getOutlookCalendarConfig());
}

function toGraphDateTime(iso: string, timeZone: string) {
  // Graph expects local-style dateTime without trailing Z when timeZone is set.
  const trimmed = iso.trim().replace(/Z$/i, "");
  return {
    dateTime: trimmed.includes(".") ? trimmed.replace(/\.\d+$/, "") : trimmed,
    timeZone,
  };
}

function buildGraphEventBody(payload: OutlookEventPayload) {
  return {
    subject: payload.subject,
    body: {
      contentType: "HTML",
      content: payload.bodyHtml,
    },
    start: toGraphDateTime(payload.startIso, payload.timeZone),
    end: toGraphDateTime(payload.endIso, payload.timeZone),
    location: payload.location
      ? { displayName: payload.location }
      : undefined,
  };
}

function mapGraphEventResult(
  data: Record<string, unknown>,
  calendarId: string | null,
): OutlookEventResult {
  const id = typeof data.id === "string" ? data.id : "";
  if (!id) {
    throw new Error("Outlook event response did not include an id.");
  }
  const iCalUId =
    (typeof data.iCalUId === "string" && data.iCalUId) ||
    (typeof data.uid === "string" && data.uid) ||
    null;
  return {
    outlookEventId: id,
    outlookCalendarId: calendarId,
    outlookICalUid: iCalUId,
  };
}

function eventsCollectionPath(config: OutlookCalendarConfig): string {
  const user = encodeURIComponent(config.userId);
  if (config.calendarId) {
    return `/users/${user}/calendars/${encodeURIComponent(config.calendarId)}/events`;
  }
  return `/users/${user}/events`;
}

function eventItemPath(
  config: OutlookCalendarConfig,
  outlookEventId: string,
): string {
  const user = encodeURIComponent(config.userId);
  return `/users/${user}/events/${encodeURIComponent(outlookEventId)}`;
}

/**
 * Creates an Outlook calendar event via Microsoft Graph.
 */
export async function createOutlookEvent(
  payload: OutlookEventPayload,
): Promise<OutlookEventResult> {
  const config = getOutlookCalendarConfig();
  if (!config) {
    throw new Error(
      "Outlook calendar is not configured. Set OUTLOOK_USER_ID (and optionally OUTLOOK_CALENDAR_ID).",
    );
  }

  const client = getGraphClient();
  const response = (await client
    .api(eventsCollectionPath(config))
    .post(buildGraphEventBody(payload))) as Record<string, unknown>;

  return mapGraphEventResult(response, config.calendarId);
}

/**
 * Updates an existing Outlook event. Does not create a duplicate.
 */
export async function updateOutlookEvent(
  outlookEventId: string,
  payload: OutlookEventPayload,
): Promise<OutlookEventResult> {
  const config = getOutlookCalendarConfig();
  if (!config) {
    throw new Error(
      "Outlook calendar is not configured. Set OUTLOOK_USER_ID (and optionally OUTLOOK_CALENDAR_ID).",
    );
  }

  const client = getGraphClient();
  const response = (await client
    .api(eventItemPath(config, outlookEventId))
    .patch(buildGraphEventBody(payload))) as Record<string, unknown>;

  // Some Graph PATCH responses omit body; fall back to known id.
  if (!response?.id) {
    return {
      outlookEventId,
      outlookCalendarId: config.calendarId,
      outlookICalUid: null,
    };
  }
  return mapGraphEventResult(response, config.calendarId);
}
