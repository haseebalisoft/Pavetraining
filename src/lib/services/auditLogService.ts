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
  AuditLogSource,
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

const NESTED_PAYLOAD_KEYS = [
  "record",
  "company",
  "permission",
  "document",
  "result",
  "settings",
] as const;

/** Fallback id keys for payloads that don't expose a plain `id` (e.g. cascade-delete summaries use `companyId`). */
const ID_FALLBACK_KEYS = [
  "companyId",
  "candidateId",
  "workforceId",
  "permissionId",
  "recordId",
  "documentId",
];

function pickId(obj: Record<string, unknown>): string | null {
  const direct = obj.id;
  if (typeof direct === "string" || typeof direct === "number") {
    return String(direct);
  }
  for (const key of ID_FALLBACK_KEYS) {
    const value = obj[key];
    if (typeof value === "string" || typeof value === "number") {
      return String(value);
    }
  }
  return null;
}

export function extractItemId(payload: unknown): string | null {
  if (!payload || typeof payload !== "object") {
    return null;
  }

  const record = payload as Record<string, unknown>;
  for (const key of NESTED_PAYLOAD_KEYS) {
    const nested = record[key];
    if (nested && typeof nested === "object") {
      const id = pickId(nested as Record<string, unknown>);
      if (id) return id;
    }
  }

  return pickId(record);
}

/** Name-ish fields checked in priority order — first match wins. */
const NAME_FALLBACK_KEYS = [
  "companyName",
  "candidateName",
  "name",
  "fullName",
  "title",
  "userEmail",
  "email",
];

function pickName(obj: Record<string, unknown>): string | null {
  for (const key of NAME_FALLBACK_KEYS) {
    const value = obj[key];
    if (typeof value === "string" && value.trim()) {
      return value.trim();
    }
  }
  return null;
}

/** Best-effort human-readable entity name from a handler's response body. */
export function extractItemName(payload: unknown): string | null {
  if (!payload || typeof payload !== "object") {
    return null;
  }

  const record = payload as Record<string, unknown>;
  for (const key of NESTED_PAYLOAD_KEYS) {
    const nested = record[key];
    if (nested && typeof nested === "object") {
      const name = pickName(nested as Record<string, unknown>);
      if (name) return name;
    }
  }

  return pickName(record);
}

type NotificationNotesPayload = {
  kind: "notification";
  type?: string;
  status?: string;
  recipientEmail?: string | null;
  companyName?: string | null;
  subject?: string | null;
  itemId?: string | null;
  errorMessage?: string | null;
  detail?: string | null;
};

function parseNotificationNotes(
  notes: string | null | undefined,
): NotificationNotesPayload | null {
  if (!notes?.trim()) return null;
  try {
    const parsed = JSON.parse(notes) as Record<string, unknown>;
    if (parsed.kind !== "notification") return null;
    return {
      kind: "notification",
      type: typeof parsed.type === "string" ? parsed.type : undefined,
      status: typeof parsed.status === "string" ? parsed.status : undefined,
      recipientEmail:
        typeof parsed.recipientEmail === "string" ? parsed.recipientEmail : null,
      companyName:
        typeof parsed.companyName === "string" ? parsed.companyName : null,
      subject: typeof parsed.subject === "string" ? parsed.subject : null,
      itemId: typeof parsed.itemId === "string" ? parsed.itemId : null,
      errorMessage:
        typeof parsed.errorMessage === "string" ? parsed.errorMessage : null,
      detail: typeof parsed.detail === "string" ? parsed.detail : null,
    };
  } catch {
    return null;
  }
}

function classifyAuditSource(input: {
  parsedAudit: AuditNotesPayload | null;
  notification: NotificationNotesPayload | null;
  roleType: string | null;
  action: string;
}): AuditLogSource {
  if (input.notification) return "notification";
  if (!input.parsedAudit) return "sharepoint";

  const action = input.action.toUpperCase();
  if (action.startsWith("NOTIFICATION_")) return "notification";
  if (action.startsWith("ADMIN_") || action === "SETTINGS_UPDATE") {
    return "admin";
  }

  const role = (input.roleType ?? "").trim().toLowerCase();
  if (role === "admin") return "admin";
  if (
    role === "customer" ||
    role === "training manager" ||
    role === "trainingmanager" ||
    role === "supervisor" ||
    role === "candidate"
  ) {
    return "customer";
  }

  if (
    action.startsWith("DOCUMENT_") ||
    action === "CANDIDATE_VIEW" ||
    action === "ACCESS_DENIED"
  ) {
    return "customer";
  }

  return "admin";
}

function mapSharePointItemToRecord(item: {
  id: string;
  fields: Record<string, unknown>;
  createdDateTime?: string | null;
}): AuditLogRecord {
  const notes = asNullableString(item.fields.Notes);
  const parsed = parseNotes(notes);
  const notification = parsed ? null : parseNotificationNotes(notes);
  const legacy = parsed || notification ? null : mapLegacyNotes(notes);

  let action =
    parsed?.action ||
    legacy?.action ||
    asString(item.fields.Area_x0020_Viewed) ||
    asString(item.fields.Title) ||
    "UNKNOWN";
  let entityType =
    parsed?.entityType ||
    asString(item.fields.ListName) ||
    "Unknown";
  let entityName = parsed?.entityName || asString(item.fields.Title) || null;
  let entityId =
    parsed?.entityId || asNullableString(item.fields.ItemsId) || null;
  let success = parsed ? parsed.success : legacy?.success !== false;
  let errorMessage = parsed?.errorMessage || legacy?.errorMessage || null;
  let company =
    parsed?.company || asNullableString(item.fields.Company) || null;
  let metadata = parsed?.metadata || null;

  if (notification) {
    const status = (notification.status ?? "unknown").toUpperCase();
    action = `NOTIFICATION_${status}`;
    entityType = "Notifications";
    entityName = notification.subject || entityName;
    entityId = notification.itemId || entityId;
    success = notification.status === "sent" || notification.status === "queued";
    errorMessage = notification.errorMessage || errorMessage;
    company = notification.companyName || company;
    metadata = sanitizeAuditMetadata({
      type: notification.type ?? null,
      status: notification.status ?? null,
      recipientEmail: notification.recipientEmail ?? null,
      detail: notification.detail ?? null,
    });
  } else if (!parsed && notes && !metadata) {
    metadata = sanitizeAuditMetadata({ notes: notes.slice(0, 300) });
  }

  const roleType =
    parsed?.roleType || asNullableString(item.fields.Role) || null;
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
      (typeof notification?.recipientEmail === "string"
        ? notification.recipientEmail.toLowerCase()
        : "") ||
      "unknown",
    roleType,
    company,
    action,
    entityType,
    entityId,
    entityName,
    success,
    errorMessage,
    ipAddress: parsed?.ipAddress || null,
    userAgent: parsed?.userAgent || null,
    metadata,
    title: asString(item.fields.Title) || action,
    source: classifyAuditSource({
      parsedAudit: parsed,
      notification,
      roleType,
      action,
    }),
  };
}

const AUDIT_LIST_FETCH_CAP = 8000;

async function fetchTrainingManagerLogItems(top: number) {
  const requested = Math.min(Math.max(top, 1), AUDIT_LIST_FETCH_CAP);
  // Small callers (dashboard) still read a window of rows so in-memory
  // newest-first sort has something recent to show.
  const cap = requested < 500 ? 500 : requested;
  // Do not use Graph `$orderby` here: SharePoint list items + `$expand=fields`
  // often hang or time out (~10s), which blocked /admin after login.
  return getListItemsByKey("trainingManagerLogs", {
    top: cap,
    skipCache: requested >= 500,
  });
}

export async function listAuditLogs(
  query: AuditLogQuery = {},
): Promise<AuditLogRecord[]> {
  // Not configured is an intentional, expected state (surfaced in the UI via
  // the "usingConsoleFallback" banner) — distinct from a genuine read failure
  // below, which is left to THROW so the API route's error handling and the
  // admin UI's error banner fire instead of silently rendering "No log
  // entries" for what is actually a SharePoint/Graph outage.
  if (!logsConfigured()) {
    return [];
  }

  try {
    const top = Math.min(Math.max(query.top ?? 2000, 1), AUDIT_LIST_FETCH_CAP);
    const items = await fetchTrainingManagerLogItems(top);
    let rows = items.map((item) =>
      mapSharePointItemToRecord({
        id: item.id,
        fields: item.fields,
        createdDateTime: item.createdDateTime,
      }),
    );

    const search = query.search?.trim().toLowerCase();
    if (search) {
      rows = rows.filter((row) => {
        const haystack = [
          row.userEmail,
          row.action,
          row.entityType,
          row.entityName,
          row.company,
          row.roleType,
          row.title,
        ]
          .filter(Boolean)
          .join(" ")
          .toLowerCase();
        return haystack.includes(search);
      });
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
    if (query.source && query.source !== "all") {
      rows = rows.filter((row) => row.source === query.source);
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

    rows.sort((a, b) => {
      const timeDiff =
        new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime();
      if (timeDiff !== 0) return timeDiff;
      return Number(b.id) - Number(a.id);
    });
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

export async function logFolderCreateFailed(input: {
  scope: "company" | "candidate";
  entityName: string;
  errorMessage: string;
  metadata?: Record<string, unknown> | null;
}): Promise<void> {
  await writeAuditEvent({
    userEmail: "system",
    roleType: "Admin",
    action: "FOLDER_CREATE_FAILED",
    entityType: "Document Folders",
    entityName: input.entityName,
    success: false,
    errorMessage: input.errorMessage,
    metadata: { scope: input.scope, ...input.metadata },
  });
}

export async function logDepartmentChange(input: {
  action:
    | "DEPARTMENT_CREATE"
    | "DEPARTMENT_UPDATE"
    | "DEPARTMENT_DEACTIVATE"
    | "DEPARTMENT_DELETE";
  userEmail: string;
  departmentId?: string | null;
  departmentName: string;
  companyName?: string | null;
  success: boolean;
  errorMessage?: string | null;
  request?: Request | null;
  metadata?: Record<string, unknown> | null;
}): Promise<void> {
  await writeAuditEvent({
    userEmail: input.userEmail,
    roleType: "Admin",
    company: input.companyName,
    action: input.action,
    entityType: "Departments",
    entityId: input.departmentId,
    entityName: input.departmentName,
    success: input.success,
    errorMessage: input.errorMessage,
    request: input.request,
    metadata: { companyName: input.companyName ?? null, ...input.metadata },
  });
}

export async function logWorkforceDepartmentAssign(input: {
  userEmail: string;
  workforceId?: string | null;
  candidateName?: string | null;
  departmentName: string;
  companyName?: string | null;
  success: boolean;
  errorMessage?: string | null;
  request?: Request | null;
  metadata?: Record<string, unknown> | null;
}): Promise<void> {
  await writeAuditEvent({
    userEmail: input.userEmail,
    roleType: "Admin",
    company: input.companyName,
    action: "WORKFORCE_DEPARTMENT_ASSIGN",
    entityType: "Workforce",
    entityId: input.workforceId,
    entityName: input.candidateName,
    success: input.success,
    errorMessage: input.errorMessage,
    request: input.request,
    metadata: {
      departmentName: input.departmentName,
      companyName: input.companyName ?? null,
      ...input.metadata,
    },
  });
}

export async function logPermissionDepartmentScopeUpdate(input: {
  userEmail: string;
  permissionId?: string | null;
  personName?: string | null;
  departmentNames: string[];
  companyName?: string | null;
  success: boolean;
  errorMessage?: string | null;
  request?: Request | null;
  metadata?: Record<string, unknown> | null;
}): Promise<void> {
  await writeAuditEvent({
    userEmail: input.userEmail,
    roleType: "Admin",
    company: input.companyName,
    action: "PERMISSION_DEPARTMENT_SCOPE_UPDATE",
    entityType: "Permissions",
    entityId: input.permissionId,
    entityName: input.personName,
    success: input.success,
    errorMessage: input.errorMessage,
    request: input.request,
    metadata: {
      departmentNames: input.departmentNames,
      companyName: input.companyName ?? null,
      ...input.metadata,
    },
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
