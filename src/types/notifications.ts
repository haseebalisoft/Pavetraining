export type NotificationType =
  | "document_upload"
  | "booking_confirmed"
  | "training_record_change"
  | "expiry_3m"
  | "expiry_6m"
  | "expiry_expired"
  | "admin_alert"
  | "portal_invite"
  /** Assigned to / removed from a candidate, or portal access removed. */
  | "access_change"
  | "login_otp"
  | "test";

/**
 * Notification kinds an Admin can switch on or off in Settings → Notification
 * types. Login codes and test emails are never switchable (sign-in and
 * diagnostics must always work). Document uploads, expiry reminders and admin
 * alerts keep their existing master switches in the same settings card.
 */
export type NotificationCategory =
  | "training_npors"
  | "training_eusr"
  | "training_streetworks"
  | "training_in_house"
  | "training_nvq"
  | "booking_confirmed"
  | "portal_invite"
  | "candidate_assignment"
  | "access_removed";

export const NOTIFICATION_CATEGORIES: ReadonlyArray<{
  key: NotificationCategory;
  label: string;
  description: string;
}> = [
  {
    key: "training_npors",
    label: "NPORS results",
    description:
      "Emails the candidate's assigned Training Manager when an NPORS record is added, updated or its Pass/Fail result is saved.",
  },
  {
    key: "training_eusr",
    label: "EUSR results",
    description: "Same, for EUSR records.",
  },
  {
    key: "training_streetworks",
    label: "Streetworks results",
    description: "Same, for Streetworks / NRSWA records.",
  },
  {
    key: "training_in_house",
    label: "In-House certificates",
    description: "Same, for In-House certificate records.",
  },
  {
    key: "training_nvq",
    label: "NVQ updates",
    description: "Same, for NVQ records.",
  },
  {
    key: "booking_confirmed",
    label: "Booking confirmations",
    description: "Emails customers when a course booking is confirmed.",
  },
  {
    key: "portal_invite",
    label: "Portal invitations",
    description: "Emails new or updated Permissions users their portal invite.",
  },
  {
    key: "candidate_assignment",
    label: "Training Manager / Supervisor assignment",
    description:
      "Emails a Training Manager or Supervisor when a candidate is assigned to or removed from them.",
  },
  {
    key: "access_removed",
    label: "Access removed",
    description: "Emails a user when their portal access is switched off.",
  },
];

export function defaultNotificationTypes(): Record<NotificationCategory, boolean> {
  return Object.fromEntries(
    NOTIFICATION_CATEGORIES.map((category) => [category.key, true]),
  ) as Record<NotificationCategory, boolean>;
}

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
