import "server-only";

import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import {
  asBoolean,
  asNullableString,
  asString,
  getListItemByKey,
  updateListItemFieldsByKey,
  type SharePointFields,
} from "@/lib/services/sharePointListService";

const eventFields = getSharePointFields("events");

/**
 * One-way sync: SharePoint / Admin Portal → Outlook.
 * Two-way sync fields are stored for later; this path never writes Outlook→SP
 * to avoid loops.
 *
 * Loop prevention:
 * - Skip when DoNotSync is true
 * - Skip when LastSyncSource is Outlook (inbound reserved for future)
 * - Mark LastSyncSource = SharePoint on outbound attempts
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

function outlookConfigured(): boolean {
  return Boolean(
    process.env.OUTLOOK_CALENDAR_ID?.trim() ||
      process.env.OUTLOOK_USER_ID?.trim(),
  );
}

/**
 * Marks an event for outbound Outlook sync and attempts push when configured.
 * Safe no-op when Outlook env is not set — leaves Pending status for ops.
 */
export async function syncEventSharePointToOutlook(
  eventId: string,
): Promise<void> {
  const item = await getListItemByKey("events", eventId);
  if (!item) {
    return;
  }

  const gate = shouldAttemptSharePointToOutlookSync(item.fields);
  if (!gate.ok) {
    if (gate.reason.toLowerCase().includes("do not sync")) {
      await updateListItemFieldsByKey("events", eventId, {
        [eventFields.syncStatus]: "Skipped",
        [eventFields.syncDirection]: "SharePointToOutlook",
        [eventFields.lastSyncSource]: "SharePoint",
        [eventFields.syncError]: gate.reason,
        [eventFields.lastSyncedAt]: new Date().toISOString(),
      });
    }
    return;
  }

  const title =
    asString(item.fields[eventFields.title]) ?? `Event ${eventId}`;

  if (!outlookConfigured()) {
    await updateListItemFieldsByKey("events", eventId, {
      [eventFields.syncStatus]: "Pending",
      [eventFields.syncDirection]: "SharePointToOutlook",
      [eventFields.lastSyncSource]: "SharePoint",
      [eventFields.syncError]:
        "Outlook sync is not configured yet. Event saved in SharePoint.",
      [eventFields.lastSyncedAt]: new Date().toISOString(),
    });
    return;
  }

  try {
    // Placeholder for Graph calendar create/update using OutlookEventId.
    // Keep SharePoint as source of truth; never pull Outlook changes here.
    await updateListItemFieldsByKey("events", eventId, {
      [eventFields.syncStatus]: "Pending",
      [eventFields.syncDirection]: "SharePointToOutlook",
      [eventFields.lastSyncSource]: "SharePoint",
      [eventFields.syncError]:
        `Outbound sync queued for "${title}" (Outlook connector pending).`,
      [eventFields.lastSyncedAt]: new Date().toISOString(),
    });
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Outlook sync failed.";
    await updateListItemFieldsByKey("events", eventId, {
      [eventFields.syncStatus]: "Failed",
      [eventFields.syncDirection]: "SharePointToOutlook",
      [eventFields.lastSyncSource]: "SharePoint",
      [eventFields.syncError]: message,
      [eventFields.lastSyncedAt]: new Date().toISOString(),
    });
  }
}
