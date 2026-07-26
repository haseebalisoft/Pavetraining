import "server-only";

import { getSharePointSiteApiRoot } from "@/lib/config/sharepoint";
import { getGraphClient } from "@/lib/graph/graphClient";
import { writeAuditLog } from "@/lib/services/auditLogService";
import { ValidationError } from "@/lib/services/validationService";
import type {
  CalendarSyncDirection,
  PortalSettings,
  PortalSettingsResponse,
  PortalSettingsSource,
  ReminderFrequency,
} from "@/types/portalSettings";

const SETTINGS_TITLE = "PAVE Portal Settings";
const SETTINGS_JSON_FIELD_CANDIDATES = [
  "SettingsJson",
  "SettingValue",
  "Notes",
  "Description",
] as const;

let cache: {
  expiresAt: number;
  payload: PortalSettingsResponse;
} | null = null;

const CACHE_MS = 30_000;

export function getDefaultSettings(): PortalSettings {
  return {
    enableCustomerNotifications: true,
    enableDocumentUploadNotifications: true,
    enableExpiryReminders: true,
    enableAdminAlerts: true,
    suppressNotificationsDuringBulkUpload: true,

    urgentWindowDays: 90,
    upcomingWindowDays: 270,
    send3MonthReminders: true,
    send6MonthReminders: true,
    sendExpiredReminders: false,
    reminderFrequency: "Weekly",

    notifyOnCertificateUpload: true,
    notifyOnCardScanUpload: true,
    notifyOnNvqDocumentUpload: true,
    requireCustomerVisibleBeforeNotification: true,
    requireNotifyCustomerBeforeNotification: true,

    allowCustomerDownloadsGlobally: true,
    showDobOnCandidateProfile: true,
    hideDobOnMatrixExceptSecondary: true,
    allowPdfSnapshotExportPlaceholder: false,
    disableExcelCsvExportForCustomers: true,

    enableOutlookSync: true,
    defaultDoNotSyncForNewTestEvents: true,
    calendarSyncDirection: "SharePointToOutlook",

    enableAuditLogging: true,
    logCustomerLogins: true,
    logDocumentViews: true,
    logDocumentDownloads: true,
    logAdminChanges: true,
    logDeniedAccessAttempts: true,
  };
}

function asBool(value: unknown, fallback: boolean): boolean {
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value === 1;
  if (typeof value === "string") {
    const normalized = value.trim().toLowerCase();
    if (["true", "yes", "1", "on"].includes(normalized)) return true;
    if (["false", "no", "0", "off"].includes(normalized)) return false;
  }
  return fallback;
}

function asPositiveInt(value: unknown, fallback: number): number {
  const num = typeof value === "number" ? value : Number(value);
  if (!Number.isFinite(num) || num <= 0) return fallback;
  return Math.floor(num);
}

function asReminderFrequency(
  value: unknown,
  fallback: ReminderFrequency,
): ReminderFrequency {
  const text = String(value ?? "").trim().toLowerCase();
  if (text === "weekly") return "Weekly";
  if (text === "monthly") return "Monthly";
  return fallback;
}

function asSyncDirection(
  value: unknown,
  fallback: CalendarSyncDirection,
): CalendarSyncDirection {
  const text = String(value ?? "").trim().toLowerCase();
  if (text === "twoway" || text === "two-way" || text === "two_way") {
    return "TwoWay";
  }
  if (
    text === "sharepointtooutlook" ||
    text === "sharepoint-to-outlook" ||
    text === "sp_to_outlook"
  ) {
    return "SharePointToOutlook";
  }
  return fallback;
}

export function mergeWithDefaults(
  partial: Partial<PortalSettings> | null | undefined,
): PortalSettings {
  const defaults = getDefaultSettings();
  if (!partial) return defaults;

  return {
    enableCustomerNotifications: asBool(
      partial.enableCustomerNotifications,
      defaults.enableCustomerNotifications,
    ),
    enableDocumentUploadNotifications: asBool(
      partial.enableDocumentUploadNotifications,
      defaults.enableDocumentUploadNotifications,
    ),
    enableExpiryReminders: asBool(
      partial.enableExpiryReminders,
      defaults.enableExpiryReminders,
    ),
    enableAdminAlerts: asBool(
      partial.enableAdminAlerts,
      defaults.enableAdminAlerts,
    ),
    suppressNotificationsDuringBulkUpload: asBool(
      partial.suppressNotificationsDuringBulkUpload,
      defaults.suppressNotificationsDuringBulkUpload,
    ),

    urgentWindowDays: asPositiveInt(
      partial.urgentWindowDays,
      defaults.urgentWindowDays,
    ),
    upcomingWindowDays: asPositiveInt(
      partial.upcomingWindowDays,
      defaults.upcomingWindowDays,
    ),
    send3MonthReminders: asBool(
      partial.send3MonthReminders,
      defaults.send3MonthReminders,
    ),
    send6MonthReminders: asBool(
      partial.send6MonthReminders,
      defaults.send6MonthReminders,
    ),
    sendExpiredReminders: asBool(
      partial.sendExpiredReminders,
      defaults.sendExpiredReminders,
    ),
    reminderFrequency: asReminderFrequency(
      partial.reminderFrequency,
      defaults.reminderFrequency,
    ),

    notifyOnCertificateUpload: asBool(
      partial.notifyOnCertificateUpload,
      defaults.notifyOnCertificateUpload,
    ),
    notifyOnCardScanUpload: asBool(
      partial.notifyOnCardScanUpload,
      defaults.notifyOnCardScanUpload,
    ),
    notifyOnNvqDocumentUpload: asBool(
      partial.notifyOnNvqDocumentUpload,
      defaults.notifyOnNvqDocumentUpload,
    ),
    requireCustomerVisibleBeforeNotification: asBool(
      partial.requireCustomerVisibleBeforeNotification,
      defaults.requireCustomerVisibleBeforeNotification,
    ),
    requireNotifyCustomerBeforeNotification: asBool(
      partial.requireNotifyCustomerBeforeNotification,
      defaults.requireNotifyCustomerBeforeNotification,
    ),

    allowCustomerDownloadsGlobally: asBool(
      partial.allowCustomerDownloadsGlobally,
      defaults.allowCustomerDownloadsGlobally,
    ),
    showDobOnCandidateProfile: asBool(
      partial.showDobOnCandidateProfile,
      defaults.showDobOnCandidateProfile,
    ),
    hideDobOnMatrixExceptSecondary: asBool(
      partial.hideDobOnMatrixExceptSecondary,
      defaults.hideDobOnMatrixExceptSecondary,
    ),
    allowPdfSnapshotExportPlaceholder: asBool(
      partial.allowPdfSnapshotExportPlaceholder,
      defaults.allowPdfSnapshotExportPlaceholder,
    ),
    disableExcelCsvExportForCustomers: asBool(
      partial.disableExcelCsvExportForCustomers,
      defaults.disableExcelCsvExportForCustomers,
    ),

    enableOutlookSync: asBool(
      partial.enableOutlookSync,
      defaults.enableOutlookSync,
    ),
    defaultDoNotSyncForNewTestEvents: asBool(
      partial.defaultDoNotSyncForNewTestEvents,
      defaults.defaultDoNotSyncForNewTestEvents,
    ),
    calendarSyncDirection: asSyncDirection(
      partial.calendarSyncDirection,
      defaults.calendarSyncDirection,
    ),

    enableAuditLogging: asBool(
      partial.enableAuditLogging,
      defaults.enableAuditLogging,
    ),
    logCustomerLogins: asBool(
      partial.logCustomerLogins,
      defaults.logCustomerLogins,
    ),
    logDocumentViews: asBool(
      partial.logDocumentViews,
      defaults.logDocumentViews,
    ),
    logDocumentDownloads: asBool(
      partial.logDocumentDownloads,
      defaults.logDocumentDownloads,
    ),
    logAdminChanges: asBool(
      partial.logAdminChanges,
      defaults.logAdminChanges,
    ),
    logDeniedAccessAttempts: asBool(
      partial.logDeniedAccessAttempts,
      defaults.logDeniedAccessAttempts,
    ),
  };
}

export function validateSettings(
  input: Partial<PortalSettings>,
): PortalSettings {
  const merged = mergeWithDefaults(input);

  if (merged.urgentWindowDays > merged.upcomingWindowDays) {
    throw new ValidationError(
      "Urgent window days must be less than or equal to upcoming window days.",
    );
  }
  if (merged.urgentWindowDays > 365 || merged.upcomingWindowDays > 730) {
    throw new ValidationError("Expiry window days are out of allowed range.");
  }

  return merged;
}

function defaultsResponse(reason?: string): PortalSettingsResponse {
  if (reason) {
    console.info("[settings]", reason);
  }
  return {
    settings: getDefaultSettings(),
    source: "defaults",
    usingDefaults: true,
    listConfigured: Boolean(getConfiguredListId()),
    updatedAt: null,
    itemId: null,
  };
}

function getConfiguredListId(): string | null {
  return process.env.SHAREPOINT_PORTAL_SETTINGS_LIST_ID?.trim() || null;
}

function extractJsonFromFields(
  fields: Record<string, unknown>,
): Partial<PortalSettings> | null {
  for (const key of SETTINGS_JSON_FIELD_CANDIDATES) {
    const value = fields[key];
    if (typeof value !== "string" || !value.trim()) continue;
    try {
      const parsed = JSON.parse(value) as unknown;
      if (parsed && typeof parsed === "object" && !Array.isArray(parsed)) {
        return parsed as Partial<PortalSettings>;
      }
    } catch {
      // try next field
    }
  }
  return null;
}

async function resolveListId(): Promise<string | null> {
  const configured = getConfiguredListId();
  if (configured) return configured;

  // Best-effort discovery by list display name — never throws to callers.
  try {
    const siteRoot = getSharePointSiteApiRoot();
    const client = getGraphClient();
    const response = (await client
      .api(`${siteRoot}/lists`)
      .filter("displayName eq 'Portal Settings'")
      .select("id,displayName")
      .top(5)
      .get()) as { value?: Array<{ id?: string; displayName?: string }> };

    const match = response.value?.find((list) =>
      list.displayName?.trim().toLowerCase() === "portal settings",
    );
    return match?.id ? String(match.id) : null;
  } catch (error) {
    console.info(
      "[settings] Portal Settings list not found; using defaults.",
      error instanceof Error ? error.message : error,
    );
    return null;
  }
}

async function readSettingsItem(listId: string): Promise<{
  id: string;
  fields: Record<string, unknown>;
  lastModifiedDateTime?: string | null;
} | null> {
  const siteRoot = getSharePointSiteApiRoot();
  const client = getGraphClient();
  const response = (await client
    .api(`${siteRoot}/lists/${listId}/items`)
    .expand("fields")
    .top(20)
    .get()) as {
    value?: Array<{
      id: string;
      lastModifiedDateTime?: string;
      fields?: Record<string, unknown>;
    }>;
  };

  const items = response.value ?? [];
  if (items.length === 0) return null;

  const preferred =
    items.find((item) => {
      const title = item.fields?.Title ?? item.fields?.title;
      return String(title ?? "").trim() === SETTINGS_TITLE;
    }) ?? items[0];

  if (!preferred) return null;
  return {
    id: String(preferred.id),
    fields: preferred.fields ?? {},
    lastModifiedDateTime: preferred.lastModifiedDateTime ?? null,
  };
}

function invalidateSettingsCache() {
  cache = null;
}

/**
 * Loads portal settings from SharePoint Portal Settings when available.
 * Falls back to defaults without crashing.
 */
export async function getSettings(
  options?: { bypassCache?: boolean },
): Promise<PortalSettingsResponse> {
  if (!options?.bypassCache && cache && cache.expiresAt > Date.now()) {
    return cache.payload;
  }

  try {
    const listId = await resolveListId();
    if (!listId) {
      const payload = defaultsResponse(
        "Using default settings (Portal Settings list not configured).",
      );
      cache = { expiresAt: Date.now() + CACHE_MS, payload };
      return payload;
    }

    const item = await readSettingsItem(listId);
    if (!item) {
      const payload: PortalSettingsResponse = {
        settings: getDefaultSettings(),
        source: "defaults",
        usingDefaults: true,
        listConfigured: true,
        updatedAt: null,
        itemId: null,
      };
      cache = { expiresAt: Date.now() + CACHE_MS, payload };
      return payload;
    }

    const parsed = extractJsonFromFields(item.fields);
    const settings = mergeWithDefaults(parsed);
    const payload: PortalSettingsResponse = {
      settings,
      source: parsed ? "sharepoint" : "defaults",
      usingDefaults: !parsed,
      listConfigured: true,
      updatedAt: item.lastModifiedDateTime ?? null,
      itemId: item.id,
    };
    cache = { expiresAt: Date.now() + CACHE_MS, payload };
    return payload;
  } catch (error) {
    console.error("[settings] Failed to load Portal Settings", error);
    const payload = defaultsResponse(
      "Using default settings (failed to read Portal Settings list).",
    );
    cache = { expiresAt: Date.now() + Math.min(CACHE_MS, 10_000), payload };
    return payload;
  }
}

async function pickWritableJsonField(listId: string): Promise<string> {
  try {
    const siteRoot = getSharePointSiteApiRoot();
    const client = getGraphClient();
    const response = (await client
      .api(`${siteRoot}/lists/${listId}/columns`)
      .select("name")
      .top(200)
      .get()) as { value?: Array<{ name?: string }> };
    const names = new Set(
      (response.value ?? [])
        .map((column) => column.name?.trim())
        .filter((name): name is string => Boolean(name)),
    );
    for (const candidate of SETTINGS_JSON_FIELD_CANDIDATES) {
      if (names.has(candidate)) return candidate;
    }
  } catch {
    // fall through
  }
  return "Notes";
}

/**
 * Persists settings to SharePoint when the list exists.
 * If the list is missing, returns defaults and does not crash.
 */
export async function updateSettings(
  input: Partial<PortalSettings>,
  options?: { actorEmail?: string | null },
): Promise<PortalSettingsResponse> {
  const validated = validateSettings(input);
  const listId = await resolveListId();

  if (!listId) {
    invalidateSettingsCache();
    await writeAuditLog({
      userEmail: options?.actorEmail ?? "unknown",
      action: "PATCH",
      entityName: "portal-settings",
      success: false,
      errorMessage:
        "Portal Settings list not available — changes not persisted (defaults only).",
    });
    return {
      settings: validated,
      source: "defaults",
      usingDefaults: true,
      listConfigured: false,
      updatedAt: null,
      itemId: null,
    };
  }

  const siteRoot = getSharePointSiteApiRoot();
  const client = getGraphClient();
  const jsonField = await pickWritableJsonField(listId);
  const payloadFields: Record<string, unknown> = {
    Title: SETTINGS_TITLE,
    [jsonField]: JSON.stringify(validated),
  };

  try {
    const existing = await readSettingsItem(listId);
    let itemId = existing?.id ?? null;

    if (itemId) {
      await client
        .api(`${siteRoot}/lists/${listId}/items/${itemId}/fields`)
        .patch(payloadFields);
    } else {
      const created = (await client.api(`${siteRoot}/lists/${listId}/items`).post({
        fields: payloadFields,
      })) as { id?: string };
      itemId = created.id ? String(created.id) : null;
    }

    invalidateSettingsCache();
    await writeAuditLog({
      userEmail: options?.actorEmail ?? "unknown",
      action: "PATCH",
      entityName: "portal-settings",
      itemId,
      success: true,
    });

    const response: PortalSettingsResponse = {
      settings: validated,
      source: "sharepoint",
      usingDefaults: false,
      listConfigured: true,
      updatedAt: new Date().toISOString(),
      itemId,
    };
    cache = { expiresAt: Date.now() + CACHE_MS, payload: response };
    return response;
  } catch (error) {
    console.error("[settings] Failed to update Portal Settings", error);
    await writeAuditLog({
      userEmail: options?.actorEmail ?? "unknown",
      action: "PATCH",
      entityName: "portal-settings",
      success: false,
      errorMessage:
        error instanceof Error ? error.message : "Failed to save settings",
    });
    throw new ValidationError(
      "Could not save settings to SharePoint. Check Portal Settings list permissions/columns.",
    );
  }
}

export type SettingsSource = PortalSettingsSource;
