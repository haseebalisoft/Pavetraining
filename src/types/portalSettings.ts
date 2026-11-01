import type { NotificationCategory } from "@/types/notifications";

export type ReminderFrequency = "Weekly" | "Monthly";

export type CalendarSyncDirection = "SharePointToOutlook" | "TwoWay";

export interface PortalSettings {
  // 1. Notification settings
  enableCustomerNotifications: boolean;
  enableDocumentUploadNotifications: boolean;
  enableExpiryReminders: boolean;
  enableAdminAlerts: boolean;
  suppressNotificationsDuringBulkUpload: boolean;
  /**
   * Per-type on/off switches (Settings → Notification types). A type that is
   * off is never sent; each skipped send is written to the notification log.
   */
  notificationTypes: Record<NotificationCategory, boolean>;

  // 2. Expiry reminder settings
  urgentWindowDays: number;
  upcomingWindowDays: number;
  send3MonthReminders: boolean;
  send6MonthReminders: boolean;
  sendExpiredReminders: boolean;
  reminderFrequency: ReminderFrequency;
  /** Workforce item ids whose expiry reminder emails are switched off. */
  expiryRemindersOffWorkforceIds: string[];

  // 3. Document settings
  notifyOnCertificateUpload: boolean;
  notifyOnCardScanUpload: boolean;
  notifyOnNvqDocumentUpload: boolean;
  requireCustomerVisibleBeforeNotification: boolean;
  requireNotifyCustomerBeforeNotification: boolean;

  // 4. Customer portal settings
  allowCustomerDownloadsGlobally: boolean;
  showDobOnCandidateProfile: boolean;
  hideDobOnMatrixExceptSecondary: boolean;
  allowPdfSnapshotExportPlaceholder: boolean;
  disableExcelCsvExportForCustomers: boolean;

  // 5. Calendar settings
  enableOutlookSync: boolean;
  defaultDoNotSyncForNewTestEvents: boolean;
  calendarSyncDirection: CalendarSyncDirection;

  // 6. Audit settings
  enableAuditLogging: boolean;
  logCustomerLogins: boolean;
  logDocumentViews: boolean;
  logDocumentDownloads: boolean;
  logAdminChanges: boolean;
  logDeniedAccessAttempts: boolean;
}

export type PortalSettingsSource = "sharepoint" | "defaults";

export interface PortalSettingsResponse {
  settings: PortalSettings;
  source: PortalSettingsSource;
  usingDefaults: boolean;
  listConfigured: boolean;
  updatedAt: string | null;
  itemId: string | null;
}
