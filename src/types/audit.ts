/**
 * Client-safe audit log types for Admin Logs UI.
 *
 * Future retention: do not auto-delete Training Manager Logs yet.
 * When adding retention, purge only after an explicit admin policy
 * (e.g. keep N days) and never from customer-facing code paths.
 */

export type AuditAction =
  | "LOGIN_SUCCESS"
  | "LOGIN_DENIED"
  | "ACCESS_DENIED"
  | "ADMIN_CREATE"
  | "ADMIN_UPDATE"
  | "ADMIN_DELETE"
  | "DOCUMENT_VIEW"
  | "DOCUMENT_DOWNLOAD"
  | "DOCUMENT_UPLOAD"
  | "CANDIDATE_VIEW"
  | "BULK_UPLOAD_PREVIEW"
  | "BULK_UPLOAD_COMMIT"
  | "NOTIFICATION_SENT"
  | "NOTIFICATION_FAILED"
  | "MATRIX_SYNC_STARTED"
  | "MATRIX_SYNC_COMPLETED"
  | "MATRIX_SYNC_FAILED"
  | "EVENT_SYNC_STARTED"
  | "EVENT_SYNC_SUCCEEDED"
  | "EVENT_SYNC_FAILED"
  | "EVENT_SYNC_SKIPPED"
  | "EVENT_SYNC_RETRY"
  | "SYSTEM_ERROR"
  | "SETTINGS_UPDATE"
  | "FOLDER_CREATE_FAILED"
  | "DEPARTMENT_CREATE"
  | "DEPARTMENT_UPDATE"
  | "DEPARTMENT_DEACTIVATE"
  | "DEPARTMENT_DELETE"
  | "WORKFORCE_DEPARTMENT_ASSIGN"
  | "PERMISSION_DEPARTMENT_SCOPE_UPDATE"
  | string;

/** Where the row originated — used to split Admin vs SharePoint/customer views. */
export type AuditLogSource =
  | "admin"
  | "customer"
  | "notification"
  | "sharepoint";

export interface AuditLogRecord {
  id: string;
  timestamp: string;
  userEmail: string;
  roleType: string | null;
  company: string | null;
  action: AuditAction;
  entityType: string;
  entityId: string | null;
  entityName: string | null;
  success: boolean;
  errorMessage: string | null;
  ipAddress: string | null;
  userAgent: string | null;
  metadata: Record<string, unknown> | null;
  title: string;
  source: AuditLogSource;
}

export interface AuditLogQuery {
  search?: string | null;
  action?: string | null;
  success?: "all" | "true" | "false" | null;
  entityType?: string | null;
  source?: AuditLogSource | "all" | null;
  from?: string | null;
  to?: string | null;
  top?: number;
}
