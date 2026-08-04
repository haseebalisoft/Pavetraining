import "server-only";

import { getSettings, getDefaultSettings } from "@/lib/services/settingsService";
import type { PortalSettings } from "@/types/portalSettings";
import type { NotificationSettingsSummary } from "@/types/notifications";

function envFlag(name: string, defaultValue: boolean): boolean {
  const raw = process.env[name]?.trim().toLowerCase();
  if (!raw) return defaultValue;
  if (["1", "true", "yes", "on"].includes(raw)) return true;
  if (["0", "false", "no", "off"].includes(raw)) return false;
  return defaultValue;
}

function envInt(name: string, defaultValue: number): number {
  const raw = process.env[name]?.trim();
  if (!raw) return defaultValue;
  const value = Number(raw);
  return Number.isFinite(value) && value > 0 ? Math.floor(value) : defaultValue;
}

/**
 * Effective notification summary for UI/logs.
 * Portal settings win for feature toggles; env still supplies mail mailbox/URL.
 */
export async function getNotificationSettings(): Promise<NotificationSettingsSummary> {
  const portal = await getSettings();
  const settings = portal.settings;
  const fromEmail = process.env.NOTIFICATION_FROM_EMAIL?.trim() || null;
  const portalUrl =
    process.env.AUTH_URL?.trim() ||
    process.env.NEXT_PUBLIC_APP_URL?.trim() ||
    null;

  // Env kill-switches still apply if explicitly set false.
  const envNotifications = envFlag("NOTIFICATIONS_ENABLED", true);
  const envExpiry = envFlag("EXPIRY_REMINDERS_ENABLED", true);

  return {
    // Master switch for customer-facing mail. Per-type flags (document,
    // booking, expiry) are enforced in sendNotification / callers.
    notificationsEnabled:
      envNotifications && settings.enableCustomerNotifications,
    expiryRemindersEnabled:
      envExpiry &&
      settings.enableCustomerNotifications &&
      settings.enableExpiryReminders,
    expiredRemindersEnabled: settings.sendExpiredReminders,
    expiry6mMaxDays: settings.upcomingWindowDays,
    fromEmail,
    emailConfigured: Boolean(fromEmail),
    portalUrl,
  };
}

/** Sync fallback for templates that only need URL/from — uses defaults + env. */
export function getNotificationSettingsSync(): NotificationSettingsSummary {
  const defaults = getDefaultSettings();
  const fromEmail = process.env.NOTIFICATION_FROM_EMAIL?.trim() || null;
  const portalUrl =
    process.env.AUTH_URL?.trim() ||
    process.env.NEXT_PUBLIC_APP_URL?.trim() ||
    null;
  return {
    notificationsEnabled: envFlag("NOTIFICATIONS_ENABLED", true),
    expiryRemindersEnabled: envFlag("EXPIRY_REMINDERS_ENABLED", true),
    expiredRemindersEnabled:
      defaults.sendExpiredReminders ||
      envFlag("EXPIRED_REMINDERS_ENABLED", false),
    expiry6mMaxDays: envInt(
      "EXPIRY_REMINDER_6M_MAX_DAYS",
      defaults.upcomingWindowDays,
    ),
    fromEmail,
    emailConfigured: Boolean(fromEmail),
    portalUrl,
  };
}

export async function getPortalSettingsCached(): Promise<PortalSettings> {
  return (await getSettings()).settings;
}

export function isNotifiableDocumentType(
  documentType: string | null | undefined,
  settings?: PortalSettings,
): boolean {
  if (!documentType?.trim()) return false;
  const cfg = settings ?? getDefaultSettings();
  const normalized = documentType.trim().toLowerCase();
  if (normalized === "certificate") return cfg.notifyOnCertificateUpload;
  if (normalized === "card scan") return cfg.notifyOnCardScanUpload;
  if (normalized === "nvq document") return cfg.notifyOnNvqDocumentUpload;
  return false;
}
