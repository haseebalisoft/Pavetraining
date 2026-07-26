export type EventSyncStatus =
  | "Pending"
  | "Synced"
  | "Failed"
  | "Skipped"
  | string;

export type EventSyncDirection =
  | "None"
  | "SharePointToOutlook"
  | "OutlookToSharePoint"
  | "TwoWay"
  | string;

export type EventSyncSource = "None" | "SharePoint" | "Outlook" | "System" | string;

export type OutlookEventPayload = {
  subject: string;
  bodyHtml: string;
  startIso: string;
  endIso: string;
  location: string | null;
  timeZone: string;
};

export type OutlookEventResult = {
  outlookEventId: string;
  outlookCalendarId: string | null;
  outlookICalUid: string | null;
};

export type CalendarSyncResult = {
  eventId: string;
  status: EventSyncStatus;
  skipped: boolean;
  reason?: string | null;
  outlookEventId?: string | null;
  error?: string | null;
};

export type EventSyncStatusSummary = {
  configured: boolean;
  enableOutlookSync: boolean;
  pending: number;
  failed: number;
  synced: number;
  skipped: number;
  total: number;
};
