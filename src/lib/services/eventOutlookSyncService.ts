import "server-only";

/**
 * Compatibility entry for admin event create/update.
 * Implementation lives in calendar/calendarSyncService (Phase 1 SharePoint → Outlook).
 */
export {
  syncEventSharePointToOutlook,
  syncEventToOutlook,
  shouldAttemptSharePointToOutlookSync,
} from "@/lib/services/calendar/calendarSyncService";

/** @deprecated Prefer syncEventToOutlook — kept for older imports. */
export { syncEventSharePointToOutlook as default } from "@/lib/services/calendar/calendarSyncService";
