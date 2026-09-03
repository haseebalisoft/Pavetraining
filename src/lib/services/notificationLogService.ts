import "server-only";

import {
  createListItemByKey,
  getListItemsByKey,
  toSharePointFields,
} from "@/lib/services/sharePointListService";
import type {
  NotificationDeliveryStatus,
  NotificationLogEntry,
  NotificationType,
} from "@/types/notifications";

const LOG_LIST_NAME = "Notifications";

export interface WriteNotificationLogInput {
  type: NotificationType;
  status: NotificationDeliveryStatus;
  recipientEmail?: string | null;
  companyName?: string | null;
  subject: string;
  dedupeKey?: string | null;
  itemId?: string | null;
  errorMessage?: string | null;
  detail?: string | null;
  actorEmail?: string | null;
}

function logsConfigured(): boolean {
  return Boolean(process.env.SHAREPOINT_TRAINING_MANAGER_LOGS_LIST_ID?.trim());
}

function encodeNotes(input: WriteNotificationLogInput): string {
  const payload = {
    kind: "notification",
    type: input.type,
    status: input.status,
    recipientEmail: input.recipientEmail ?? null,
    companyName: input.companyName ?? null,
    subject: input.subject,
    dedupeKey: input.dedupeKey ?? null,
    itemId: input.itemId ?? null,
    errorMessage: input.errorMessage ?? null,
    detail: input.detail ?? null,
  };
  return JSON.stringify(payload);
}

function parseNotes(notes: string | null | undefined): Partial<WriteNotificationLogInput> | null {
  if (!notes?.trim()) return null;
  try {
    const parsed = JSON.parse(notes) as Record<string, unknown>;
    if (parsed.kind !== "notification") return null;
    return {
      type: parsed.type as NotificationType,
      status: parsed.status as NotificationDeliveryStatus,
      recipientEmail: (parsed.recipientEmail as string | null) ?? null,
      companyName: (parsed.companyName as string | null) ?? null,
      subject: (parsed.subject as string) ?? "",
      dedupeKey: (parsed.dedupeKey as string | null) ?? null,
      itemId: (parsed.itemId as string | null) ?? null,
      errorMessage: (parsed.errorMessage as string | null) ?? null,
      detail: (parsed.detail as string | null) ?? null,
    };
  } catch {
    return null;
  }
}

/**
 * Persists a notification attempt to Training Manager Logs (when configured).
 * Failures are swallowed so logging never breaks delivery flow.
 */
export async function writeNotificationLog(
  input: WriteNotificationLogInput,
): Promise<string | null> {
  const timestamp = new Date().toISOString();
  const title = `Notify · ${input.type} · ${input.status}${
    input.recipientEmail ? ` · ${input.recipientEmail}` : ""
  }`;

  console.info("[notification-log]", {
    ...input,
    timestamp,
  });

  if (!logsConfigured()) {
    return null;
  }

  try {
    const fields = toSharePointFields("trainingManagerLogs", {
      title: title.slice(0, 240),
      userEmail: (input.actorEmail ?? "system").trim().toLowerCase(),
      listName: LOG_LIST_NAME,
      itemsId: input.itemId ?? null,
      areaViewed: LOG_LIST_NAME,
      timestamp,
      notes: encodeNotes(input),
      company: input.companyName ?? null,
    });
    const item = await createListItemByKey("trainingManagerLogs", fields);
    return item.id;
  } catch (error) {
    console.error("[notification-log] Failed to write SharePoint log", error);
    return null;
  }
}

export async function listNotificationLogs(options?: {
  top?: number;
  failedOnly?: boolean;
}): Promise<NotificationLogEntry[]> {
  const top = options?.top ?? 100;
  if (!logsConfigured()) {
    return [];
  }

  try {
    // Shared, growing list — SharePoint's default order is oldest-id-first,
    // so a `top` cap without `orderBy` would silently drop the newest rows
    // (and with them, recent sends the dedupe check needs to see).
    const items = await getListItemsByKey("trainingManagerLogs", {
      top: 2000,
      orderBy: "id desc",
    });
    const mapped: NotificationLogEntry[] = [];

    for (const item of items) {
      const notes =
        typeof item.fields.Notes === "string" ? item.fields.Notes : null;
      const parsed = parseNotes(notes);
      if (!parsed?.type || !parsed.status) continue;
      if (options?.failedOnly && parsed.status !== "failed") continue;

      const createdAt =
        (typeof item.fields.Timestamp === "string" && item.fields.Timestamp) ||
        item.createdDateTime ||
        new Date().toISOString();

      mapped.push({
        id: item.id,
        type: parsed.type,
        status: parsed.status,
        recipientEmail: parsed.recipientEmail ?? null,
        companyName: parsed.companyName ?? null,
        subject: parsed.subject || "Notification",
        dedupeKey: parsed.dedupeKey ?? null,
        itemId: parsed.itemId ?? null,
        errorMessage: parsed.errorMessage ?? null,
        createdAt,
        notes: parsed.detail ?? null,
      });
    }

    mapped.sort(
      (a, b) =>
        new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime(),
    );
    return mapped.slice(0, top);
  } catch (error) {
    console.error("[notification-log] Failed to list logs", error);
    return [];
  }
}

/**
 * True when a successful/queued notification with this dedupe key already exists
 * inside the lookback window (default 30 days).
 *
 * Failed / skipped / not_configured do NOT block retries — otherwise fixing
 * mail config or recipient prefs still leaves bookings silently undelivered.
 */
export async function hasRecentNotificationDedupe(
  dedupeKey: string,
  lookbackDays = 30,
): Promise<boolean> {
  if (!dedupeKey.trim()) return false;
  const logs = await listNotificationLogs({ top: 500 });
  const cutoff = Date.now() - lookbackDays * 86_400_000;
  return logs.some((entry) => {
    if (entry.dedupeKey !== dedupeKey) return false;
    if (entry.status !== "sent" && entry.status !== "queued") {
      return false;
    }
    return new Date(entry.createdAt).getTime() >= cutoff;
  });
}
