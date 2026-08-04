import "server-only";

import { getGraphClient } from "@/lib/graph/graphClient";
import type {
  OutlookEventPayload,
  OutlookEventResult,
} from "@/types/calendarSync";

export type OutlookCalendarConfig =
  | {
      mode: "user";
      userId: string;
      calendarId: string | null;
    }
  | {
      mode: "group";
      groupId: string;
      calendarId: string | null;
    };

/**
 * Outlook mailbox / group calendar config for app-only Graph access.
 *
 * Prefer group calendar for shared ops:
 *   OUTLOOK_GROUP_ID = Microsoft 365 group id (e.g. Pave Training Operations)
 * Optional:
 *   OUTLOOK_CALENDAR_ID = specific calendar id; omit for default group calendar
 *
 * Fallback (shared mailbox / user):
 *   OUTLOOK_USER_ID = mailbox UPN or Graph user id
 */
export function getOutlookCalendarConfig(): OutlookCalendarConfig | null {
  const groupId = process.env.OUTLOOK_GROUP_ID?.trim() || "";
  const userId = process.env.OUTLOOK_USER_ID?.trim() || "";
  const calendarId = process.env.OUTLOOK_CALENDAR_ID?.trim() || null;

  if (groupId) {
    return { mode: "group", groupId, calendarId };
  }
  if (userId) {
    return { mode: "user", userId, calendarId };
  }
  return null;
}

export function isOutlookCalendarConfigured(): boolean {
  return Boolean(getOutlookCalendarConfig());
}

export function describeOutlookCalendarTarget(): string {
  const config = getOutlookCalendarConfig();
  if (!config) return "not configured";
  if (config.mode === "group") {
    return config.calendarId
      ? `group:${config.groupId}/calendar:${config.calendarId}`
      : `group:${config.groupId}/default`;
  }
  return config.calendarId
    ? `user:${config.userId}/calendar:${config.calendarId}`
    : `user:${config.userId}/default`;
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
    showAs: payload.showAs ?? "busy",
    categories: ["PAVE Training Portal"],
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
  if (config.mode === "group") {
    const group = encodeURIComponent(config.groupId);
    if (config.calendarId) {
      return `/groups/${group}/calendars/${encodeURIComponent(config.calendarId)}/events`;
    }
    return `/groups/${group}/events`;
  }
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
  if (config.mode === "group") {
    const group = encodeURIComponent(config.groupId);
    return `/groups/${group}/events/${encodeURIComponent(outlookEventId)}`;
  }
  const user = encodeURIComponent(config.userId);
  return `/users/${user}/events/${encodeURIComponent(outlookEventId)}`;
}

function configMissingMessage(): string {
  return "Outlook calendar is not configured. Set OUTLOOK_GROUP_ID (Pave Training Operations) or OUTLOOK_USER_ID, and optionally OUTLOOK_CALENDAR_ID.";
}

/**
 * Creates an Outlook calendar event via Microsoft Graph.
 */
export async function createOutlookEvent(
  payload: OutlookEventPayload,
): Promise<OutlookEventResult> {
  const config = getOutlookCalendarConfig();
  if (!config) {
    throw new Error(configMissingMessage());
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
    throw new Error(configMissingMessage());
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

/**
 * Deletes an Outlook calendar event. No-ops when Outlook is not configured.
 * Treats already-missing events as success.
 */
export async function deleteOutlookEvent(
  outlookEventId: string,
): Promise<void> {
  const config = getOutlookCalendarConfig();
  if (!config || !outlookEventId.trim()) return;

  const client = getGraphClient();
  try {
    await client.api(eventItemPath(config, outlookEventId)).delete();
  } catch (error) {
    const status =
      typeof error === "object" &&
      error &&
      "statusCode" in error &&
      typeof (error as { statusCode?: unknown }).statusCode === "number"
        ? (error as { statusCode: number }).statusCode
        : null;
    if (status === 404) return;
    throw error;
  }
}
