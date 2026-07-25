import "server-only";

import {
  createListItemByKey,
  toSharePointFields,
} from "@/lib/services/sharePointListService";

export interface AuditLogInput {
  userEmail: string;
  action: string;
  entityName: string;
  itemId?: string | null;
  success: boolean;
  errorMessage?: string | null;
}

function logsConfigured(): boolean {
  return Boolean(process.env.SHAREPOINT_TRAINING_MANAGER_LOGS_LIST_ID?.trim());
}

/**
 * Writes an audit entry to Training Manager Logs when configured.
 * Failures are swallowed so logging never breaks the primary request.
 */
export async function writeAuditLog(input: AuditLogInput): Promise<void> {
  const timestamp = new Date().toISOString();
  const entry = {
    ...input,
    userEmail: input.userEmail.trim().toLowerCase(),
    timestamp,
  };

  if (!logsConfigured()) {
    console.info("[audit]", entry);
    return;
  }

  try {
    const title = `${entry.action} · ${entry.entityName}${
      entry.itemId ? ` · ${entry.itemId}` : ""
    }`;

    const fields = toSharePointFields("trainingManagerLogs", {
      title,
      userEmail: entry.userEmail,
      action: entry.action,
      entityName: entry.entityName,
      itemId: entry.itemId ?? null,
      timestamp: entry.timestamp,
      success: entry.success,
      errorMessage: entry.errorMessage ?? null,
    });

    await createListItemByKey("trainingManagerLogs", fields);
  } catch (error) {
    console.error("[audit] Failed to write Training Manager Logs entry", error);
    console.info("[audit:fallback]", entry);
  }
}

export function extractItemId(payload: unknown): string | null {
  if (!payload || typeof payload !== "object") {
    return null;
  }

  const record = payload as Record<string, unknown>;
  for (const key of ["record", "company", "permission", "document"]) {
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
