export type NotificationType =
  | "document_upload"
  | "expiry_3m"
  | "expiry_6m"
  | "expiry_expired"
  | "admin_alert"
  | "test";

export type NotificationDeliveryStatus =
  | "sent"
  | "queued"
  | "skipped"
  | "failed"
  | "not_configured";

export interface NotificationSettingsSummary {
  notificationsEnabled: boolean;
  expiryRemindersEnabled: boolean;
  expiredRemindersEnabled: boolean;
  expiry6mMaxDays: number;
  fromEmail: string | null;
  emailConfigured: boolean;
  portalUrl: string | null;
}

export interface NotificationLogEntry {
  id: string;
  type: NotificationType;
  status: NotificationDeliveryStatus;
  recipientEmail: string | null;
  companyName: string | null;
  subject: string;
  dedupeKey: string | null;
  itemId: string | null;
  errorMessage: string | null;
  createdAt: string;
  notes: string | null;
}

export interface NotificationSendResult {
  status: NotificationDeliveryStatus;
  recipientEmail: string;
  subject: string;
  messageId?: string | null;
  errorMessage?: string | null;
  logged: boolean;
}

export interface DocumentNotificationResult {
  documentId: string;
  attempted: boolean;
  skipped: boolean;
  skipReason?: string;
  recipients: string[];
  results: NotificationSendResult[];
  notificationSent: boolean;
  errorMessage?: string | null;
}

export interface ExpiryNotificationRunResult {
  dryRun: boolean;
  companiesProcessed: number;
  remindersPrepared: number;
  emailsAttempted: number;
  sent: number;
  skipped: number;
  failed: number;
  notConfigured: number;
  items: Array<{
    companyName: string;
    recipientEmail: string;
    window: "3m" | "6m" | "expired";
    candidateCount: number;
    status: NotificationDeliveryStatus;
    errorMessage?: string | null;
  }>;
}
