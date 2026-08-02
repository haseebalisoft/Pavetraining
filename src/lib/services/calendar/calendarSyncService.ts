import "server-only";

import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import { writeAuditEvent } from "@/lib/services/auditLogService";
import { bookingStatusFromFreeBusy } from "@/lib/services/bookingStatusService";
import { computeEventSyncHash } from "@/lib/services/calendar/eventHashService";
import { mapSharePointEventToOutlookPayload } from "@/lib/services/calendar/eventMappingService";
import {
  createOutlookEvent,
  isOutlookCalendarConfigured,
  updateOutlookEvent,
} from "@/lib/services/calendar/outlookCalendarService";
import { getCompanyById } from "@/lib/services/companyService";
import { getPortalSettingsCached } from "@/lib/services/notificationConfig";
import {
  asBoolean,
  asLookupOrString,
  asNullableString,
  asString,
  getListItemByKey,
  getListItemsByKey,
  listHasColumn,
  updateListItemFieldsByKey,
  type SharePointFields,
  type SharePointListItem,
} from "@/lib/services/sharePointListService";
import { stripSharePointHtml } from "@/lib/text/stripSharePointHtml";
import type {
  CalendarSyncResult,
  EventSyncStatusSummary,
  OutlookEventResult,
} from "@/types/calendarSync";

const eventFields = getSharePointFields("events");

const SYNC_ERROR_MAX = 255;

/**
 * Gate used by older call sites / tests. Prefer syncEventToOutlook which
 * applies the same rules plus settings, hash skip, and Outlook API.
 */
export function shouldAttemptSharePointToOutlookSync(
  fields: SharePointFields,
): { ok: true } | { ok: false; reason: string } {
  if (asBoolean(fields[eventFields.doNotSync])) {
    return { ok: false, reason: "Do not sync is enabled for this event." };
  }

  const source = (asNullableString(fields[eventFields.lastSyncSource]) ?? "")
    .trim()
    .toLowerCase();
  if (source === "outlook") {
    return {
      ok: false,
      reason: "Skipped to prevent sync loop (last source was Outlook).",
    };
  }

  return { ok: true };
}

export type SyncEventOptions = {
  /** Bypass SyncHash skip (and Outlook last-source loop skip). */
  force?: boolean;
  userEmail?: string | null;
  request?: Request | null;
  /** When true, clears DoNotSync before syncing (admin "enable sync" flows). */
  clearDoNotSync?: boolean;
};

function truncateSyncError(message: string | null | undefined): string | null {
  if (!message?.trim()) return null;
  const text = message.trim();
  if (text.length <= SYNC_ERROR_MAX) return text;
  return `${text.slice(0, SYNC_ERROR_MAX - 1)}…`;
}

function safePublicError(error: unknown): string {
  if (error instanceof Error && error.message.trim()) {
    const message = error.message.trim();
    // Strip common Graph noise / secrets-ish fragments.
    return message
      .replace(/Bearer\s+\S+/gi, "[redacted]")
      .replace(/client_secret[=:]\S+/gi, "[redacted]")
      .slice(0, 500);
  }
  return "Outlook sync failed.";
}

function resolveCompanyId(fields: SharePointFields): string | null {
  return (
    asString(fields[eventFields.eventCompanyLookupId]) ??
    asString(fields.EventCompanyId) ??
    (typeof fields[eventFields.eventCompany] === "object"
      ? asString(
          (fields[eventFields.eventCompany] as { LookupId?: unknown }).LookupId,
        )
      : null) ??
    null
  );
}

async function resolveCompanyName(
  fields: SharePointFields,
  companyId: string | null,
): Promise<string | null> {
  const fromLookup = asLookupOrString(fields[eventFields.eventCompany]);
  if (fromLookup?.trim()) return fromLookup.trim();
  if (!companyId) return null;
  try {
    const company = await getCompanyById(companyId);
    return company?.companyName?.trim() || null;
  } catch {
    return null;
  }
}

function readEventCore(fields: SharePointFields) {
  return {
    title: asString(fields[eventFields.title]) ?? "Untitled event",
    eventDate: asNullableString(fields[eventFields.eventDate]),
    endDate: asNullableString(fields[eventFields.endDate]),
    location: asNullableString(fields[eventFields.location]),
    description: stripSharePointHtml(
      asNullableString(fields[eventFields.description]),
    ),
    trainingAddress: stripSharePointHtml(
      asNullableString(fields[eventFields.trainingAddress]),
    ),
    bookingStatus: bookingStatusFromFreeBusy(fields[eventFields.freeBusy]),
    outlookEventId: asNullableString(fields[eventFields.outlookEventId]),
    syncStatus: asNullableString(fields[eventFields.syncStatus]),
    syncHash: asNullableString(fields[eventFields.syncHash]),
    lastSyncSource: asNullableString(fields[eventFields.lastSyncSource]),
    doNotSync: asBoolean(fields[eventFields.doNotSync]),
  };
}

async function optionalSyncColumns(): Promise<{
  hasLastSyncedAt: boolean;
  hasSyncHash: boolean;
  hasOutlookCalendarId: boolean;
  hasOutlookICalUid: boolean;
}> {
  const [hasLastSyncedAt, hasSyncHash, hasOutlookCalendarId, hasOutlookICalUid] =
    await Promise.all([
      listHasColumn("events", eventFields.lastSyncedAt),
      listHasColumn("events", eventFields.syncHash),
      listHasColumn("events", eventFields.outlookCalendarId),
      listHasColumn("events", eventFields.outlookICalUid),
    ]);
  return {
    hasLastSyncedAt,
    hasSyncHash,
    hasOutlookCalendarId,
    hasOutlookICalUid,
  };
}

async function auditSync(input: {
  action: string;
  eventId: string;
  title: string;
  success: boolean;
  errorMessage?: string | null;
  userEmail?: string | null;
  request?: Request | null;
  metadata?: Record<string, unknown> | null;
}): Promise<void> {
  await writeAuditEvent({
    userEmail: input.userEmail ?? "system",
    roleType: "Admin",
    action: input.action,
    entityType: "Events",
    entityId: input.eventId,
    entityName: input.title,
    success: input.success,
    errorMessage: input.errorMessage ?? null,
    request: input.request,
    metadata: input.metadata ?? null,
  });
}

/**
 * Writes sync metadata fields that exist on the live Events list.
 * Never throws — callers handle failures.
 */
export async function skipSync(
  eventId: string,
  reason: string,
): Promise<void> {
  await updateListItemFieldsByKey("events", eventId, {
    [eventFields.syncStatus]: "Skipped",
    [eventFields.syncDirection]: "SharePointToOutlook",
    [eventFields.lastSyncSource]: "SharePoint",
    [eventFields.syncError]: truncateSyncError(reason),
  });
}

export async function markSyncFailure(
  eventId: string,
  error: unknown,
): Promise<string> {
  const publicError = truncateSyncError(safePublicError(error)) ?? "Outlook sync failed.";
  await updateListItemFieldsByKey("events", eventId, {
    [eventFields.syncStatus]: "Failed",
    [eventFields.syncDirection]: "SharePointToOutlook",
    [eventFields.lastSyncSource]: "SharePoint",
    [eventFields.syncError]: publicError,
  });
  return publicError;
}

export async function markSyncSuccess(
  eventId: string,
  outlookResult: OutlookEventResult,
  syncHash: string,
): Promise<void> {
  const cols = await optionalSyncColumns();
  const fields: SharePointFields = {
    [eventFields.syncStatus]: "Synced",
    [eventFields.syncDirection]: "SharePointToOutlook",
    [eventFields.lastSyncSource]: "SharePoint",
    [eventFields.syncError]: "",
    [eventFields.outlookEventId]: outlookResult.outlookEventId,
  };

  if (cols.hasOutlookCalendarId && outlookResult.outlookCalendarId) {
    fields[eventFields.outlookCalendarId] = outlookResult.outlookCalendarId;
  }
  if (cols.hasOutlookICalUid && outlookResult.outlookICalUid) {
    fields[eventFields.outlookICalUid] = outlookResult.outlookICalUid;
  }
  if (cols.hasLastSyncedAt) {
    fields[eventFields.lastSyncedAt] = new Date().toISOString();
  }
  if (cols.hasSyncHash) {
    fields[eventFields.syncHash] = syncHash;
  }

  await updateListItemFieldsByKey("events", eventId, fields);
}

export async function createOutlookEventFromSharePointEvent(
  item: SharePointListItem,
): Promise<{ result: OutlookEventResult; syncHash: string }> {
  const core = readEventCore(item.fields);
  const companyId = resolveCompanyId(item.fields);
  const companyName = await resolveCompanyName(item.fields, companyId);
  const syncHash = computeEventSyncHash({
    title: core.title,
    eventDate: core.eventDate,
    endDate: core.endDate,
    location: core.location,
    description: core.description,
    trainingAddress: core.trainingAddress,
    companyName,
    companyId,
    bookingStatus: core.bookingStatus,
  });
  const payload = mapSharePointEventToOutlookPayload({
    title: core.title,
    eventDate: core.eventDate,
    endDate: core.endDate,
    location: core.location,
    description: core.description,
    trainingAddress: core.trainingAddress,
    companyName,
    bookingStatus: core.bookingStatus,
  });
  const result = await createOutlookEvent(payload);
  return { result, syncHash };
}

export async function updateOutlookEventFromSharePointEvent(
  item: SharePointListItem,
  outlookEventId: string,
): Promise<{ result: OutlookEventResult; syncHash: string }> {
  const core = readEventCore(item.fields);
  const companyId = resolveCompanyId(item.fields);
  const companyName = await resolveCompanyName(item.fields, companyId);
  const syncHash = computeEventSyncHash({
    title: core.title,
    eventDate: core.eventDate,
    endDate: core.endDate,
    location: core.location,
    description: core.description,
    trainingAddress: core.trainingAddress,
    companyName,
    companyId,
    bookingStatus: core.bookingStatus,
  });
  const payload = mapSharePointEventToOutlookPayload({
    title: core.title,
    eventDate: core.eventDate,
    endDate: core.endDate,
    location: core.location,
    description: core.description,
    trainingAddress: core.trainingAddress,
    companyName,
    bookingStatus: core.bookingStatus,
  });
  const result = await updateOutlookEvent(outlookEventId, payload);
  return { result, syncHash };
}

/**
 * Phase 1: SharePoint Events → Outlook Calendar (one-way).
 * SharePoint remains source of truth. Never pulls Outlook→SharePoint.
 */
export async function syncEventToOutlook(
  eventId: string,
  options: SyncEventOptions = {},
): Promise<CalendarSyncResult> {
  const userEmail = options.userEmail ?? "system";

  try {
    if (options.clearDoNotSync) {
      await updateListItemFieldsByKey("events", eventId, {
        [eventFields.doNotSync]: "No",
        [eventFields.syncStatus]: "Pending",
        [eventFields.syncDirection]: "SharePointToOutlook",
        [eventFields.lastSyncSource]: "SharePoint",
        [eventFields.syncError]: "",
      });
    }

    const item = await getListItemByKey("events", eventId);
    if (!item) {
      return {
        eventId,
        status: "Failed",
        skipped: false,
        error: "Event not found in SharePoint.",
      };
    }

    const core = readEventCore(item.fields);

    await auditSync({
      action: "EVENT_SYNC_STARTED",
      eventId,
      title: core.title,
      success: true,
      userEmail,
      request: options.request,
      metadata: { force: Boolean(options.force) },
    });

    if (core.doNotSync) {
      const reason = "Do not sync is enabled for this event.";
      await skipSync(eventId, reason);
      await auditSync({
        action: "EVENT_SYNC_SKIPPED",
        eventId,
        title: core.title,
        success: true,
        userEmail,
        request: options.request,
        metadata: { reason },
      });
      return {
        eventId,
        status: "Skipped",
        skipped: true,
        reason,
      };
    }

    const settings = await getPortalSettingsCached();
    if (!settings.enableOutlookSync) {
      const reason = "Outlook sync is disabled in portal settings.";
      await skipSync(eventId, reason);
      await auditSync({
        action: "EVENT_SYNC_SKIPPED",
        eventId,
        title: core.title,
        success: true,
        userEmail,
        request: options.request,
        metadata: { reason },
      });
      return {
        eventId,
        status: "Skipped",
        skipped: true,
        reason,
      };
    }

    // Loop prevention prep for future Outlook→SharePoint inbound.
    const source = (core.lastSyncSource ?? "").trim().toLowerCase();
    if (!options.force && source === "outlook") {
      const reason =
        "Skipped to prevent sync loop (last source was Outlook).";
      await skipSync(eventId, reason);
      await auditSync({
        action: "EVENT_SYNC_SKIPPED",
        eventId,
        title: core.title,
        success: true,
        userEmail,
        request: options.request,
        metadata: { reason },
      });
      return {
        eventId,
        status: "Skipped",
        skipped: true,
        reason,
      };
    }

    const companyId = resolveCompanyId(item.fields);
    const companyName = await resolveCompanyName(item.fields, companyId);
    const syncHash = computeEventSyncHash({
      title: core.title,
      eventDate: core.eventDate,
      endDate: core.endDate,
      location: core.location,
      description: core.description,
      trainingAddress: core.trainingAddress,
      companyName,
      companyId,
    });

    if (
      !options.force &&
      core.syncStatus?.toLowerCase() === "synced" &&
      core.syncHash &&
      core.syncHash === syncHash
    ) {
      const reason = "No changes since last successful sync.";
      await auditSync({
        action: "EVENT_SYNC_SKIPPED",
        eventId,
        title: core.title,
        success: true,
        userEmail,
        request: options.request,
        metadata: { reason, syncHash },
      });
      return {
        eventId,
        status: "Synced",
        skipped: true,
        reason,
        outlookEventId: core.outlookEventId,
      };
    }

    if (!isOutlookCalendarConfigured()) {
      const reason =
        "Outlook calendar is not configured. Set OUTLOOK_GROUP_ID (Pave Training Operations) or OUTLOOK_USER_ID, and optionally OUTLOOK_CALENDAR_ID.";
      const publicError = await markSyncFailure(eventId, new Error(reason));
      console.warn(`[calendarSync] ${eventId}: ${reason}`);
      await auditSync({
        action: "EVENT_SYNC_FAILED",
        eventId,
        title: core.title,
        success: false,
        errorMessage: publicError,
        userEmail,
        request: options.request,
      });
      return {
        eventId,
        status: "Failed",
        skipped: false,
        error: publicError,
      };
    }

    try {
      let outlookResult: OutlookEventResult;
      let hash = syncHash;

      if (core.outlookEventId) {
        const updated = await updateOutlookEventFromSharePointEvent(
          item,
          core.outlookEventId,
        );
        outlookResult = updated.result;
        hash = updated.syncHash;
      } else {
        const created = await createOutlookEventFromSharePointEvent(item);
        outlookResult = created.result;
        hash = created.syncHash;
      }

      await markSyncSuccess(eventId, outlookResult, hash);
      await auditSync({
        action: "EVENT_SYNC_SUCCEEDED",
        eventId,
        title: core.title,
        success: true,
        userEmail,
        request: options.request,
        metadata: {
          outlookEventId: outlookResult.outlookEventId,
          updated: Boolean(core.outlookEventId),
        },
      });

      return {
        eventId,
        status: "Synced",
        skipped: false,
        outlookEventId: outlookResult.outlookEventId,
      };
    } catch (error) {
      console.error(`[calendarSync] Outlook API failed for ${eventId}:`, error);
      const publicError = await markSyncFailure(eventId, error);
      await auditSync({
        action: "EVENT_SYNC_FAILED",
        eventId,
        title: core.title,
        success: false,
        errorMessage: publicError,
        userEmail,
        request: options.request,
      });
      return {
        eventId,
        status: "Failed",
        skipped: false,
        error: publicError,
      };
    }
  } catch (error) {
    console.error(`[calendarSync] unexpected error for ${eventId}:`, error);
    const publicError = truncateSyncError(safePublicError(error));
    try {
      await markSyncFailure(eventId, error);
    } catch {
      // Ignore secondary metadata failures.
    }
    await auditSync({
      action: "EVENT_SYNC_FAILED",
      eventId,
      title: `Event ${eventId}`,
      success: false,
      errorMessage: publicError,
      userEmail,
      request: options.request,
    });
    return {
      eventId,
      status: "Failed",
      skipped: false,
      error: publicError,
    };
  }
}

/**
 * Compatibility wrapper used by admin event create/update.
 * Never throws — event CRUD must succeed even if sync fails.
 */
export async function syncEventSharePointToOutlook(
  eventId: string,
): Promise<void> {
  try {
    await syncEventToOutlook(eventId);
  } catch (error) {
    console.warn(
      `[events sync] failed for ${eventId}:`,
      error instanceof Error ? error.message : error,
    );
  }
}

export async function syncPendingEvents(options: {
  userEmail?: string | null;
  request?: Request | null;
  limit?: number;
}): Promise<{
  processed: number;
  results: CalendarSyncResult[];
}> {
  const limit = Math.min(Math.max(options.limit ?? 50, 1), 200);
  const items = await getListItemsByKey("events", { top: 5000 });
  const pending = items
    .filter((item) => {
      const status = (
        asNullableString(item.fields[eventFields.syncStatus]) ?? ""
      )
        .trim()
        .toLowerCase();
      const doNotSync = asBoolean(item.fields[eventFields.doNotSync]);
      return !doNotSync && (status === "pending" || status === "failed");
    })
    .slice(0, limit);

  const results: CalendarSyncResult[] = [];
  for (const item of pending) {
    results.push(
      await syncEventToOutlook(item.id, {
        force: true,
        userEmail: options.userEmail,
        request: options.request,
      }),
    );
  }

  return { processed: results.length, results };
}

export async function getEventSyncStatusSummary(): Promise<EventSyncStatusSummary> {
  const [settings, configured, items] = await Promise.all([
    getPortalSettingsCached(),
    Promise.resolve(isOutlookCalendarConfigured()),
    getListItemsByKey("events", { top: 5000 }),
  ]);

  let pending = 0;
  let failed = 0;
  let synced = 0;
  let skipped = 0;

  for (const item of items) {
    const status = (
      asNullableString(item.fields[eventFields.syncStatus]) ?? ""
    )
      .trim()
      .toLowerCase();
    if (status === "pending") pending += 1;
    else if (status === "failed") failed += 1;
    else if (status === "synced") synced += 1;
    else if (status === "skipped") skipped += 1;
  }

  return {
    configured,
    enableOutlookSync: settings.enableOutlookSync,
    pending,
    failed,
    synced,
    skipped,
    total: items.length,
  };
}
