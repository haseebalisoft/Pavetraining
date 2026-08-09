import "server-only";

/**
 * Training Manager Logs audit writer/reader.
 *
 * Future retention: do not auto-delete old logs yet. When retention is added,
 * enforce an explicit admin policy (e.g. keep 365 days) server-side only.
 */

import {
  asNullableString,
  asString,
  createListItemByKey,
  getListItemByKey,
  getListItemsByKey,
  toSharePointFields,
} from "@/lib/services/sharePointListService";
import type {
  AuditAction,
  AuditLogQuery,
  AuditLogRecord,
} from "@/types/audit";

/** @deprecated Prefer structured log* helpers; kept for existing call sites. */
export interface AuditLogInput {
  userEmail: string;
  action: string;
  entityName: string;
  itemId?: string | null;
  success: boolean;
  errorMessage?: string | null;
  roleType?: string | null;
  company?: string | null;
  entityType?: string | null;
  entityId?: string | null;
  entityDisplayName?: string | null;
  ipAddress?: string | null;
  userAgent?: string | null;
  metadata?: Record<string, unknown> | null;
  request?: Request | null;
}

export interface AuditEventInput {
  userEmail?: string | null;
  roleType?: string | null;
  company?: string | null;
  action: AuditAction;
  entityType: string;
  entityId?: string | null;
  entityName?: string | null;
  success: boolean;
  errorMessage?: string | null;
  ipAddress?: string | null;
  userAgent?: string | null;
  metadata?: Record<string, unknown> | null;
  request?: Request | null;
}

const SECRET_KEY_PATTERN =
  /(secret|token|password|authorization|api[_-]?key|client[_-]?secret|refresh|bearer|cookie|set-cookie)/i;

function logsConfigured(): boolean {
  return Boolean(process.env.SHAREPOINT_TRAINING_MANAGER_LOGS_LIST_ID?.trim());
}

/** Strip secrets and truncate for SharePoint / UI-safe storage. */
export function sanitizeAuditError(error: unknown): string | null {
  if (!error) return null;
  let message =
    typeof error === "string"
      ? error
      : error instanceof Error
        ? error.message
        : "Unexpected error";

  message = message
    .replace(/Bearer\s+[A-Za-z0-9\-._~+/]+=*/gi, "Bearer [REDACTED]")
    .replace(
      /(client_secret|clientSecret|access_token|refresh_token|password)\s*[:=]\s*\S+/gi,
      "$1=[REDACTED]",
    )
    .replace(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g, (email) => {
      // Keep emails (needed for audit identity) — only scrub tokens above.
      return email;
    });

  // Never expose Graph request ids / correlation blobs wholesale.
  if (message.length > 400) {
    message = `${message.slice(0, 400)}…`;
  }
  return message || null;
}

export function sanitizeAuditMetadata(
  metadata: Record<string, unknown> | null | undefined,
): Record<string, unknown> | null {
  if (!metadata) return null;
  const safe: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(metadata)) {
    if (SECRET_KEY_PATTERN.test(key)) continue;
    if (value === undefined) continue;
    if (typeof value === "string") {
      if (SECRET_KEY_PATTERN.test(value)) continue;
      safe[key] = value.length > 300 ? `${value.slice(0, 300)}…` : value;
      continue;
    }
    if (
      typeof value === "number" ||
      typeof value === "boolean" ||
      value === null
    ) {
      safe[key] = value;
      continue;
    }
    if (Array.isArray(value)) {
      safe[key] = value.slice(0, 20).map((entry) =>
        typeof entry === "string" ||
        typeof entry === "number" ||
        typeof entry === "boolean" ||
        entry === null
          ? entry
          : "[object]",
      );
      continue;
    }
    // Skip nested objects that may hold secrets.
  }
  return Object.keys(safe).length > 0 ? safe : null;
}

export function requestAuditContext(request?: Request | null): {
  ipAddress: string | null;
  userAgent: string | null;
} {
  if (!request) return { ipAddress: null, userAgent: null };
  const forwarded = request.headers.get("x-forwarded-for");
  const ipAddress =
    forwarded?.split(",")[0]?.trim() ||
    request.headers.get("x-real-ip")?.trim() ||
    null;
  const userAgent = request.headers.get("user-agent")?.trim() || null;
  return { ipAddress, userAgent };
}

type AuditNotesPayload = {
  kind: "audit";
  action: string;
  entityType: string;
  entityId: string | null;
  entityName: string | null;
  success: boolean;
  errorMessage: string | null;
  roleType: string | null;
  company: string | null;
  ipAddress: string | null;
  userAgent: string | null;
  metadata: Record<string, unknown> | null;
  timestamp: string;
};

function encodeNotes(payload: AuditNotesPayload): string {
  return JSON.stringify(payload);
}

function parseNotes(notes: string | null | undefined): AuditNotesPayload | null {
  if (!notes?.trim()) return null;
  try {
    const parsed = JSON.parse(notes) as Record<string, unknown>;
    if (parsed.kind !== "audit") return null;
    return {
      kind: "audit",
      action: String(parsed.action ?? ""),
      entityType: String(parsed.entityType ?? ""),
      entityId:
        parsed.entityId === null || parsed.entityId === undefined
          ? null
          : String(parsed.entityId),
      entityName:
        parsed.entityName === null || parsed.entityName === undefined
          ? null
          : String(parsed.entityName),
      success: Boolean(parsed.success),
      errorMessage:
        typeof parsed.errorMessage === "string" ? parsed.errorMessage : null,
      roleType:
        typeof parsed.roleType === "string" ? parsed.roleType : null,
      company: typeof parsed.company === "string" ? parsed.company : null,
      ipAddress:
        typeof parsed.ipAddress === "string" ? parsed.ipAddress : null,
      userAgent:
        typeof parsed.userAgent === "string" ? parsed.userAgent : null,
      metadata:
        parsed.metadata && typeof parsed.metadata === "object"
          ? (parsed.metadata as Record<string, unknown>)
          : null,
      timestamp:
        typeof parsed.timestamp === "string"
          ? parsed.timestamp
          : new Date().toISOString(),
    };
  } catch {
    return null;
  }
}

function mapLegacyNotes(notes: string | null | undefined): Partial<AuditNotesPayload> {
  if (!notes?.trim()) return {};
  const actionMatch = notes.match(/Action:\s*(.+)/i);
  const successMatch = notes.match(/Success:\s*(Yes|No)/i);
  const errorMatch = notes.match(/Error:\s*(.+)/i);
  return {
    action: actionMatch?.[1]?.trim() ?? "",
    success: successMatch ? successMatch[1]?.toLowerCase() === "yes" : true,
    errorMessage: errorMatch?.[1]?.trim() ?? null,
  };
}

/**
 * Core writer — never throws to callers.
 */
export async function writeAuditEvent(input: AuditEventInput): Promise<void> {
  const reqCtx = requestAuditContext(input.request);
  const timestamp = new Date().toISOString();
  const userEmail = (input.userEmail ?? "unknown").trim().toLowerCase() || "unknown";
  const errorMessage = sanitizeAuditError(input.errorMessage);
  const metadata = sanitizeAuditMetadata(input.metadata);
  const entityType = input.entityType.trim() || "Unknown";
  const action = String(input.action || "UNKNOWN");

  const notesPayload: AuditNotesPayload = {
    kind: "audit",
    action,
    entityType,
    entityId: input.entityId ?? null,
    entityName: input.entityName ?? null,
    success: input.success,
    errorMessage,
    roleType: input.roleType ?? null,
    company: input.company ?? null,
    ipAddress: input.ipAddress ?? reqCtx.ipAddress,
    userAgent: input.userAgent ?? reqCtx.userAgent,
    metadata,
    timestamp,
  };

  const title = `${action} · ${entityType}${
    input.entityId ? ` · ${input.entityId}` : ""
  }${input.success ? "" : " · FAILED"}`.slice(0, 240);

  const entry = {
    ...notesPayload,
    userEmail,
    title,
  };

  if (!logsConfigured()) {
    console.info("[audit]", entry);
    return;
  }

  try {
    const fields = toSharePointFields("trainingManagerLogs", {
      title,
      userEmail,
      listName: entityType,
      itemsId: input.entityId ?? null,
      areaViewed: action,
      timestamp,
      notes: encodeNotes(notesPayload),
      company: input.company ?? null,
      role: input.roleType ?? null,
    });

    await createListItemByKey("trainingManagerLogs", fields);
  } catch (error) {
    console.error("[audit] Failed to write Training Manager Logs entry", error);
    console.info("[audit:fallback]", entry);
  }
}

/** Backward-compatible wrapper used by apiGuards and older call sites. */
export async function writeAuditLog(input: AuditLogInput): Promise<void> {
  await writeAuditEvent({
    userEmail: input.userEmail,
    roleType: input.roleType,
    company: input.company,
    action: input.action,
    entityType: input.entityType ?? input.entityName,
    entityId: input.entityId ?? input.itemId,
    entityName: input.entityDisplayName ?? input.entityName,
    success: input.success,
    errorMessage: input.errorMessage,
    ipAddress: input.ipAddress,
    userAgent: input.userAgent,
    metadata: input.metadata,
    request: input.request,
  });
}

export function extractItemId(payload: unknown): string | null {
  if (!payload || typeof payload !== "object") {
    return null;
  }

  const record = payload as Record<string, unknown>;
  for (const key of [
    "record",
    "company",
    "permission",
    "document",
    "result",
    "settings",
  ]) {
    const nested = record[key];
    if (nested && typeof nested === "object" && "id" in nested) {
      const id = (nested as { id?: unknown }).id;
      if (typeof id === "string" || typeof id === "number") {
        return String(id);
      }
    }
  }

  if (typeof record.id === "string" || typeof record.id === "number") {
    return String(record.id);
  }

  return null;
}

function mapSharePointItemToRecord(item: {
  id: string;
  fields: Record<string, unknown>;
  createdDateTime?: string | null;
}): AuditLogRecord | null {
  const notes = asNullableString(item.fields.Notes);
  const parsed = parseNotes(notes);
  const legacy = parsed ? null : mapLegacyNotes(notes);

  const action =
    parsed?.action ||
    legacy?.action ||
    asString(item.fields.Area_x0020_Viewed) ||
    asString(item.fields.Title) ||
    "UNKNOWN";

  // Skip pure notification-kind rows from notificationLogService in main audit UI
  // unless they also carry audit kind (notification logs use kind:notification).
  if (notes?.includes('"kind":"notification"') && !parsed) {
    // Still show notification failures/success as system events if desired —
    // include them with entityType Notifications.
  }

  const timestamp =
    parsed?.timestamp ||
    asNullableString(item.fields.Timestamp) ||
    item.createdDateTime ||
    new Date().toISOString();

  return {
    id: item.id,
    timestamp,
    userEmail:
      asString(item.fields.User_x0020_Email)?.toLowerCase() ||
      asString(item.fields.UserEmail)?.toLowerCase() ||
      "unknown",
    roleType:
      parsed?.roleType || asNullableString(item.fields.Role) || null,
    company:
      parsed?.company || asNullableString(item.fields.Company) || null,
    action,
    entityType:
      parsed?.entityType ||
      asString(item.fields.ListName) ||
      "Unknown",
    entityId:
      parsed?.entityId || asNullableString(item.fields.ItemsId) || null,
    entityName: parsed?.entityName || null,
    success: parsed ? parsed.success : legacy?.success !== false,
    errorMessage: parsed?.errorMessage || legacy?.errorMessage || null,
    ipAddress: parsed?.ipAddress || null,
    userAgent: parsed?.userAgent || null,
    metadata: parsed?.metadata || null,
    title: asString(item.fields.Title) || action,
  };
}

export async function listAuditLogs(
  query: AuditLogQuery = {},
): Promise<AuditLogRecord[]> {
  if (!logsConfigured()) {
    return [];
  }

  try {
    const top = Math.min(Math.max(query.top ?? 200, 1), 500);
    const items = await getListItemsByKey("trainingManagerLogs", { top: 500 });
    let rows = items
      .map((item) =>
        mapSharePointItemToRecord({
          id: item.id,
          fields: item.fields,
          createdDateTime: item.createdDateTime,
        }),
      )
      .filter((row): row is AuditLogRecord => Boolean(row));

    // Prefer structured audit + compatible legacy rows; keep notification logs too.
    const search = query.search?.trim().toLowerCase();
    if (search) {
      rows = rows.filter((row) => row.userEmail.includes(search));
    }
    if (query.action?.trim()) {
      const action = query.action.trim().toLowerCase();
      rows = rows.filter((row) => row.action.toLowerCase().includes(action));
    }
    if (query.entityType?.trim()) {
      const entityType = query.entityType.trim().toLowerCase();
      rows = rows.filter((row) =>
        row.entityType.toLowerCase().includes(entityType),
      );
    }
    if (query.success === "true") {
      rows = rows.filter((row) => row.success);
    } else if (query.success === "false") {
      rows = rows.filter((row) => !row.success);
    }
    if (query.from) {
      const fromMs = new Date(query.from).getTime();
      if (!Number.isNaN(fromMs)) {
        rows = rows.filter((row) => new Date(row.timestamp).getTime() >= fromMs);
      }
    }
    if (query.to) {
      const toDate = new Date(query.to);
      if (!Number.isNaN(toDate.getTime())) {
        toDate.setHours(23, 59, 59, 999);
        const toMs = toDate.getTime();
        rows = rows.filter((row) => new Date(row.timestamp).getTime() <= toMs);
      }
    }

    rows.sort(
      (a, b) =>
        new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime(),
    );
    return rows.slice(0, top);
  } catch (error) {
    console.error("[audit] Failed to list logs", error);
    throw new Error(
      error instanceof Error
        ? `Failed to load audit logs from SharePoint: ${error.message}`
        : "Failed to load audit logs from SharePoint.",
    );
  }
}

export async function getAuditLogById(
  logId: string,
): Promise<AuditLogRecord | null> {
  if (!logsConfigured()) return null;
  try {
    const item = await getListItemByKey("trainingManagerLogs", logId);
    if (!item) return null;
    return mapSharePointItemToRecord({
      id: item.id,
      fields: item.fields,
      createdDateTime: item.createdDateTime,
    });
  } catch (error) {
    console.error("[audit] Failed to load log", error);
    return null;
  }
}

/* ── Typed helpers ─────────────────────────────────────────── */

export async function logLogin(input: {
  userEmail: string;
  roleType?: string | null;
  company?: string | null;
  success: boolean;
  errorMessage?: string | null;
  request?: Request | null;
}): Promise<void> {
  await writeAuditEvent({
    userEmail: input.userEmail,
    roleType: input.roleType,
    company: input.company,
    action: input.success ? "LOGIN_SUCCESS" : "LOGIN_DENIED",
    entityType: "Auth",
    entityName: "Portal login",
    success: input.success,
    errorMessage: input.errorMessage,
    request: input.request,
  });
}

export async function logAccessDenied(input: {
  userEmail?: string | null;
  roleType?: string | null;
  company?: string | null;
  entityType: string;
  entityId?: string | null;
  entityName?: string | null;
  errorMessage?: string | null;
  request?: Request | null;
  metadata?: Record<string, unknown> | null;
}): Promise<void> {
  await writeAuditEvent({
    ...input,
    action: "ACCESS_DENIED",
    success: false,
    errorMessage: input.errorMessage ?? "Access denied",
  });
}

export async function logDocumentView(input: {
  userEmail: string;
  roleType?: string | null;
  company?: string | null;
  documentId: string;
  documentName?: string | null;
  success: boolean;
  errorMessage?: string | null;
  request?: Request | null;
}): Promise<void> {
  await writeAuditEvent({
    userEmail: input.userEmail,
    roleType: input.roleType,
    company: input.company,
    action: "DOCUMENT_VIEW",
    entityType: "Customer Documents",
    entityId: input.documentId,
    entityName: input.documentName ?? null,
    success: input.success,
    errorMessage: input.errorMessage,
    request: input.request,
  });
}

export async function logDocumentDownload(input: {
  userEmail: string;
  roleType?: string | null;
  company?: string | null;
  documentId: string;
  documentName?: string | null;
  success: boolean;
  errorMessage?: string | null;
  request?: Request | null;
}): Promise<void> {
  await writeAuditEvent({
    userEmail: input.userEmail,
    roleType: input.roleType,
    company: input.company,
    action: "DOCUMENT_DOWNLOAD",
    entityType: "Customer Documents",
    entityId: input.documentId,
    entityName: input.documentName ?? null,
    success: input.success,
    errorMessage: input.errorMessage,
    request: input.request,
  });
}

export async function logAdminCreate(input: {
  userEmail: string;
  entityType: string;
  entityId?: string | null;
  entityName?: string | null;
  company?: string | null;
  success?: boolean;
  errorMessage?: string | null;
  request?: Request | null;
  metadata?: Record<string, unknown> | null;
}): Promise<void> {
  await writeAuditEvent({
    userEmail: input.userEmail,
    roleType: "Admin",
    company: input.company,
    action: "ADMIN_CREATE",
    entityType: input.entityType,
    entityId: input.entityId,
    entityName: input.entityName,
    success: input.success ?? true,
    errorMessage: input.errorMessage,
    request: input.request,
    metadata: input.metadata,
  });
}

export async function logAdminUpdate(input: {
  userEmail: string;
  entityType: string;
  entityId?: string | null;
  entityName?: string | null;
  company?: string | null;
  success?: boolean;
  errorMessage?: string | null;
  request?: Request | null;
  metadata?: Record<string, unknown> | null;
}): Promise<void> {
  await writeAuditEvent({
    userEmail: input.userEmail,
    roleType: "Admin",
    company: input.company,
    action: "ADMIN_UPDATE",
    entityType: input.entityType,
    entityId: input.entityId,
    entityName: input.entityName,
    success: input.success ?? true,
    errorMessage: input.errorMessage,
    request: input.request,
    metadata: input.metadata,
  });
}

export async function logAdminDelete(input: {
  userEmail: string;
  entityType: string;
  entityId?: string | null;
  entityName?: string | null;
  company?: string | null;
  success?: boolean;
  errorMessage?: string | null;
  request?: Request | null;
}): Promise<void> {
  await writeAuditEvent({
    userEmail: input.userEmail,
    roleType: "Admin",
    company: input.company,
    action: "ADMIN_DELETE",
    entityType: input.entityType,
    entityId: input.entityId,
    entityName: input.entityName,
    success: input.success ?? true,
    errorMessage: input.errorMessage,
    request: input.request,
  });
}

export async function logBulkUpload(input: {
  userEmail: string;
  phase: "preview" | "commit";
  success: boolean;
  itemCount?: number;
  errorMessage?: string | null;
  request?: Request | null;
  metadata?: Record<string, unknown> | null;
}): Promise<void> {
  await writeAuditEvent({
    userEmail: input.userEmail,
    roleType: "Admin",
    action:
      input.phase === "preview" ? "BULK_UPLOAD_PREVIEW" : "BULK_UPLOAD_COMMIT",
    entityType: "Bulk Upload",
    success: input.success,
    errorMessage: input.errorMessage,
    request: input.request,
    metadata: {
      itemCount: input.itemCount ?? null,
      ...input.metadata,
    },
  });
}

export async function logNotification(input: {
  userEmail?: string | null;
  recipientEmail?: string | null;
  success: boolean;
  company?: string | null;
  entityId?: string | null;
  entityName?: string | null;
  errorMessage?: string | null;
  metadata?: Record<string, unknown> | null;
}): Promise<void> {
  await writeAuditEvent({
    userEmail: input.userEmail ?? "system",
    action: input.success ? "NOTIFICATION_SENT" : "NOTIFICATION_FAILED",
    entityType: "Notifications",
    entityId: input.entityId,
    entityName: input.entityName ?? input.recipientEmail ?? null,
    company: input.company,
    success: input.success,
    errorMessage: input.errorMessage,
    metadata: {
      recipientEmail: input.recipientEmail ?? null,
      ...input.metadata,
    },
  });
}

export async function logMatrixSync(input: {
  userEmail?: string | null;
  phase: "started" | "completed" | "failed";
  scope?: string | null;
  success?: boolean;
  errorMessage?: string | null;
  metadata?: Record<string, unknown> | null;
}): Promise<void> {
  const action =
    input.phase === "started"
      ? "MATRIX_SYNC_STARTED"
      : input.phase === "failed"
        ? "MATRIX_SYNC_FAILED"
        : "MATRIX_SYNC_COMPLETED";
  await writeAuditEvent({
    userEmail: input.userEmail ?? "system",
    roleType: "Admin",
    action,
    entityType: "Training Matrix",
    entityName: input.scope ?? "matrix-sync",
    success:
      input.success ??
      (input.phase === "completed"
        ? true
        : input.phase === "failed"
          ? false
          : true),
    errorMessage: input.errorMessage,
    metadata: input.metadata,
  });
}

export async function logSystemError(input: {
  userEmail?: string | null;
  entityType?: string | null;
  entityId?: string | null;
  errorMessage: string;
  metadata?: Record<string, unknown> | null;
}): Promise<void> {
  await writeAuditEvent({
    userEmail: input.userEmail ?? "system",
    action: "SYSTEM_ERROR",
    entityType: input.entityType ?? "System",
    entityId: input.entityId,
    success: false,
    errorMessage: input.errorMessage,
    metadata: input.metadata,
  });
}

export async function logCandidateView(input: {
  userEmail: string;
  roleType?: string | null;
  company?: string | null;
  candidateId: string;
  candidateName?: string | null;
  success: boolean;
  errorMessage?: string | null;
}): Promise<void> {
  await writeAuditEvent({
    userEmail: input.userEmail,
    roleType: input.roleType,
    company: input.company,
    action: input.success ? "CANDIDATE_VIEW" : "ACCESS_DENIED",
    entityType: "Workforce",
    entityId: input.candidateId,
    entityName: input.candidateName ?? null,
    success: input.success,
    errorMessage: input.errorMessage,
  });
}
