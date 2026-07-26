"use client";

import { useCallback, useEffect, useState, type ReactNode } from "react";

import { useAdminToast } from "@/components/admin/AdminToast";
import { Breadcrumbs } from "@/components/ui/Breadcrumbs";
import { LoadingState } from "@/components/ui/States";
import { readPublicApiError } from "@/lib/errors/publicMessages";
import type {
  CalendarSyncDirection,
  PortalSettings,
  PortalSettingsResponse,
  ReminderFrequency,
} from "@/types/portalSettings";

import styles from "@/components/admin/admin.module.css";

type BoolKey = {
  [K in keyof PortalSettings]: PortalSettings[K] extends boolean ? K : never;
}[keyof PortalSettings];

function ToggleRow({
  label,
  description,
  checked,
  disabled,
  onChange,
}: {
  label: string;
  description?: string;
  checked: boolean;
  disabled?: boolean;
  onChange: (value: boolean) => void;
}) {
  return (
    <label className={styles.settingsToggleRow}>
      <div>
        <span className={styles.settingsToggleLabel}>{label}</span>
        {description ? (
          <p className={styles.settingsToggleHint}>{description}</p>
        ) : null}
      </div>
      <button
        type="button"
        role="switch"
        aria-checked={checked}
        className={`${styles.settingsSwitch} ${checked ? styles.settingsSwitchOn : ""}`}
        disabled={disabled}
        onClick={() => onChange(!checked)}
      >
        <span className={styles.settingsSwitchThumb} />
      </button>
    </label>
  );
}

function SettingsSection({
  title,
  description,
  children,
}: {
  title: string;
  description: string;
  children: ReactNode;
}) {
  return (
    <section className={styles.settingsCard}>
      <header className={styles.settingsCardHeader}>
        <h2>{title}</h2>
        <p>{description}</p>
      </header>
      <div className={styles.settingsCardBody}>{children}</div>
    </section>
  );
}

export function AdminSettingsClient() {
  const { pushToast } = useAdminToast();
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [payload, setPayload] = useState<PortalSettingsResponse | null>(null);
  const [draft, setDraft] = useState<PortalSettings | null>(null);

  const load = useCallback(async () => {
    setLoading(true);
    try {
      const response = await fetch("/api/admin/settings", { cache: "no-store" });
      if (!response.ok) throw new Error(await readPublicApiError(response));
      const data = (await response.json()) as PortalSettingsResponse;
      setPayload(data);
      setDraft(data.settings);
    } catch (error) {
      pushToast(
        error instanceof Error ? error.message : "Failed to load settings",
        "error",
      );
    } finally {
      setLoading(false);
    }
  }, [pushToast]);

  useEffect(() => {
    void load();
  }, [load]);

  function setBool(key: BoolKey, value: boolean) {
    setDraft((current) => (current ? { ...current, [key]: value } : current));
  }

  function setNumber(
    key: "urgentWindowDays" | "upcomingWindowDays",
    value: number,
  ) {
    setDraft((current) =>
      current
        ? { ...current, [key]: Number.isFinite(value) ? value : current[key] }
        : current,
    );
  }

  async function save(next?: PortalSettings, resetToDefaults = false) {
    if (!draft && !next && !resetToDefaults) return;
    setSaving(true);
    try {
      const response = await fetch("/api/admin/settings", {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(
          resetToDefaults
            ? { resetToDefaults: true }
            : { settings: next ?? draft },
        ),
      });
      if (!response.ok) throw new Error(await readPublicApiError(response));
      const data = (await response.json()) as PortalSettingsResponse;
      setPayload(data);
      setDraft(data.settings);
      pushToast(
        data.listConfigured
          ? data.usingDefaults
            ? "Settings updated (list empty — defaults applied)."
            : "Settings saved."
          : "Using default settings — Portal Settings list not available to persist.",
        "success",
      );
    } catch (error) {
      pushToast(
        error instanceof Error ? error.message : "Failed to save settings",
        "error",
      );
    } finally {
      setSaving(false);
    }
  }

  if (loading || !draft) {
    return <LoadingState label="Loading settings…" />;
  }

  return (
    <div>
      <header className={styles.pageHeader}>
        <div>
          <Breadcrumbs
            items={[
              { label: "Admin", href: "/admin" },
              { label: "Settings" },
            ]}
          />
          <p className={styles.eyebrow}>Admin</p>
          <h1 className={styles.title}>Portal settings</h1>
          <p className={styles.subtitle}>
            Control notifications, expiry reminders, documents, customer portal,
            calendar, and audit behaviour without changing code.
          </p>
        </div>
        <div className={styles.settingsActions}>
          <button
            type="button"
            className={styles.secondaryButton}
            disabled={saving}
            onClick={() => void save(undefined, true)}
          >
            Reset to defaults
          </button>
          <button
            type="button"
            className={styles.primaryButton}
            disabled={saving}
            onClick={() => void save()}
          >
            {saving ? "Saving…" : "Save settings"}
          </button>
        </div>
      </header>

      <div
        className={`${styles.settingsBanner} ${
          payload?.usingDefaults
            ? styles.settingsBannerWarn
            : styles.settingsBannerOk
        }`}
        role="status"
      >
        {payload?.usingDefaults
          ? "Using default settings"
          : "Loaded from SharePoint Portal Settings"}
        {payload?.listConfigured
          ? " · list configured"
          : " · Portal Settings list not found"}
        {payload?.updatedAt
          ? ` · updated ${new Date(payload.updatedAt).toLocaleString()}`
          : ""}
      </div>

      <div className={styles.settingsLayout}>
        <SettingsSection
          title="Notification settings"
          description="Master switches for customer and admin notifications."
        >
          <ToggleRow
            label="Enable customer notifications"
            checked={draft.enableCustomerNotifications}
            onChange={(v) => setBool("enableCustomerNotifications", v)}
            disabled={saving}
          />
          <ToggleRow
            label="Enable document upload notifications"
            checked={draft.enableDocumentUploadNotifications}
            onChange={(v) => setBool("enableDocumentUploadNotifications", v)}
            disabled={saving}
          />
          <ToggleRow
            label="Enable expiry reminders"
            checked={draft.enableExpiryReminders}
            onChange={(v) => setBool("enableExpiryReminders", v)}
            disabled={saving}
          />
          <ToggleRow
            label="Enable admin alerts"
            checked={draft.enableAdminAlerts}
            onChange={(v) => setBool("enableAdminAlerts", v)}
            disabled={saving}
          />
          <ToggleRow
            label="Suppress notifications during bulk upload"
            description="When on, bulk document updates do not email customers."
            checked={draft.suppressNotificationsDuringBulkUpload}
            onChange={(v) =>
              setBool("suppressNotificationsDuringBulkUpload", v)
            }
            disabled={saving}
          />
        </SettingsSection>

        <SettingsSection
          title="Expiry reminder settings"
          description="Windows and reminder cadence for Training Matrix expiries."
        >
          <div className={styles.settingsFieldGrid}>
            <label className={styles.field}>
              <span className={styles.fieldLabel}>Urgent window days</span>
              <input
                className={styles.input}
                type="number"
                min={1}
                max={365}
                value={draft.urgentWindowDays}
                disabled={saving}
                onChange={(event) =>
                  setNumber("urgentWindowDays", Number(event.target.value))
                }
              />
            </label>
            <label className={styles.field}>
              <span className={styles.fieldLabel}>Upcoming window days</span>
              <input
                className={styles.input}
                type="number"
                min={1}
                max={730}
                value={draft.upcomingWindowDays}
                disabled={saving}
                onChange={(event) =>
                  setNumber("upcomingWindowDays", Number(event.target.value))
                }
              />
            </label>
            <label className={styles.field}>
              <span className={styles.fieldLabel}>Reminder frequency</span>
              <select
                className={styles.select}
                value={draft.reminderFrequency}
                disabled={saving}
                onChange={(event) =>
                  setDraft((current) =>
                    current
                      ? {
                          ...current,
                          reminderFrequency: event.target
                            .value as ReminderFrequency,
                        }
                      : current,
                  )
                }
              >
                <option value="Weekly">Weekly</option>
                <option value="Monthly">Monthly</option>
              </select>
            </label>
          </div>
          <ToggleRow
            label="Send 3-month reminders"
            checked={draft.send3MonthReminders}
            onChange={(v) => setBool("send3MonthReminders", v)}
            disabled={saving}
          />
          <ToggleRow
            label="Send 6-month reminders"
            checked={draft.send6MonthReminders}
            onChange={(v) => setBool("send6MonthReminders", v)}
            disabled={saving}
          />
          <ToggleRow
            label="Send expired reminders"
            checked={draft.sendExpiredReminders}
            onChange={(v) => setBool("sendExpiredReminders", v)}
            disabled={saving}
          />
        </SettingsSection>

        <SettingsSection
          title="Document settings"
          description="Which uploads can notify customers and what flags are required."
        >
          <ToggleRow
            label="Notify on Certificate upload"
            checked={draft.notifyOnCertificateUpload}
            onChange={(v) => setBool("notifyOnCertificateUpload", v)}
            disabled={saving}
          />
          <ToggleRow
            label="Notify on Card Scan upload"
            checked={draft.notifyOnCardScanUpload}
            onChange={(v) => setBool("notifyOnCardScanUpload", v)}
            disabled={saving}
          />
          <ToggleRow
            label="Notify on NVQ Document upload"
            checked={draft.notifyOnNvqDocumentUpload}
            onChange={(v) => setBool("notifyOnNvqDocumentUpload", v)}
            disabled={saving}
          />
          <ToggleRow
            label="Require CustomerVisible before notification"
            checked={draft.requireCustomerVisibleBeforeNotification}
            onChange={(v) =>
              setBool("requireCustomerVisibleBeforeNotification", v)
            }
            disabled={saving}
          />
          <ToggleRow
            label="Require NotifyCustomer before notification"
            checked={draft.requireNotifyCustomerBeforeNotification}
            onChange={(v) =>
              setBool("requireNotifyCustomerBeforeNotification", v)
            }
            disabled={saving}
          />
        </SettingsSection>

        <SettingsSection
          title="Customer portal settings"
          description="Customer-facing visibility and export behaviour."
        >
          <ToggleRow
            label="Allow customer downloads globally"
            checked={draft.allowCustomerDownloadsGlobally}
            onChange={(v) => setBool("allowCustomerDownloadsGlobally", v)}
            disabled={saving}
          />
          <ToggleRow
            label="Show DOB on candidate profile"
            checked={draft.showDobOnCandidateProfile}
            onChange={(v) => setBool("showDobOnCandidateProfile", v)}
            disabled={saving}
          />
          <ToggleRow
            label="Hide DOB on matrix except secondary text"
            checked={draft.hideDobOnMatrixExceptSecondary}
            onChange={(v) => setBool("hideDobOnMatrixExceptSecondary", v)}
            disabled={saving}
          />
          <ToggleRow
            label="Allow PDF/snapshot export (placeholder)"
            checked={draft.allowPdfSnapshotExportPlaceholder}
            onChange={(v) => setBool("allowPdfSnapshotExportPlaceholder", v)}
            disabled={saving}
          />
          <ToggleRow
            label="Disable Excel/CSV export for customers"
            checked={draft.disableExcelCsvExportForCustomers}
            onChange={(v) => setBool("disableExcelCsvExportForCustomers", v)}
            disabled={saving}
          />
        </SettingsSection>

        <SettingsSection
          title="Calendar settings"
          description="Outlook sync behaviour for Events."
        >
          <ToggleRow
            label="Enable Outlook sync"
            checked={draft.enableOutlookSync}
            onChange={(v) => setBool("enableOutlookSync", v)}
            disabled={saving}
          />
          <ToggleRow
            label="Default DoNotSync for new test events"
            checked={draft.defaultDoNotSyncForNewTestEvents}
            onChange={(v) => setBool("defaultDoNotSyncForNewTestEvents", v)}
            disabled={saving}
          />
          <label className={styles.field}>
            <span className={styles.fieldLabel}>Calendar sync direction</span>
            <select
              className={styles.select}
              value={draft.calendarSyncDirection}
              disabled={saving}
              onChange={(event) =>
                setDraft((current) =>
                  current
                    ? {
                        ...current,
                        calendarSyncDirection: event.target
                          .value as CalendarSyncDirection,
                      }
                    : current,
                )
              }
            >
              <option value="SharePointToOutlook">SharePoint → Outlook</option>
              <option value="TwoWay">Two-way (later)</option>
            </select>
          </label>
        </SettingsSection>

        <SettingsSection
          title="Audit settings"
          description="What the portal records in Training Manager Logs."
        >
          <ToggleRow
            label="Enable audit logging"
            checked={draft.enableAuditLogging}
            onChange={(v) => setBool("enableAuditLogging", v)}
            disabled={saving}
          />
          <ToggleRow
            label="Log customer logins"
            checked={draft.logCustomerLogins}
            onChange={(v) => setBool("logCustomerLogins", v)}
            disabled={saving}
          />
          <ToggleRow
            label="Log document views"
            checked={draft.logDocumentViews}
            onChange={(v) => setBool("logDocumentViews", v)}
            disabled={saving}
          />
          <ToggleRow
            label="Log document downloads"
            checked={draft.logDocumentDownloads}
            onChange={(v) => setBool("logDocumentDownloads", v)}
            disabled={saving}
          />
          <ToggleRow
            label="Log admin changes"
            checked={draft.logAdminChanges}
            onChange={(v) => setBool("logAdminChanges", v)}
            disabled={saving}
          />
          <ToggleRow
            label="Log denied access attempts"
            checked={draft.logDeniedAccessAttempts}
            onChange={(v) => setBool("logDeniedAccessAttempts", v)}
            disabled={saving}
          />
        </SettingsSection>
      </div>
    </div>
  );
}
