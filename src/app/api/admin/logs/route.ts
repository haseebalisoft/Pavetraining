import { withAdminApi } from "@/lib/api/adminApi";
import { listAuditLogs } from "@/lib/services/auditLogService";
import type { AuditLogQuery } from "@/types/audit";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  const url = new URL(request.url);
  const sourceParam = url.searchParams.get("source");
  const query: AuditLogQuery = {
    search: url.searchParams.get("search"),
    action: url.searchParams.get("action"),
    success: (url.searchParams.get("success") as AuditLogQuery["success"]) ?? "all",
    entityType: url.searchParams.get("entityType"),
    source:
      sourceParam === "admin" ||
      sourceParam === "customer" ||
      sourceParam === "notification" ||
      sourceParam === "sharepoint"
        ? sourceParam
        : "all",
    from: url.searchParams.get("from"),
    to: url.searchParams.get("to"),
    top: Number(url.searchParams.get("top") ?? "8000"),
  };

  return withAdminApi(
    "GET /api/admin/logs",
    async () => {
      const logs = await listAuditLogs(query);
      const configured = Boolean(
        process.env.SHAREPOINT_TRAINING_MANAGER_LOGS_LIST_ID?.trim(),
      );
      return {
        logs,
        total: logs.length,
        fetchedAt: new Date().toISOString(),
        configured,
        usingConsoleFallback: !configured,
        // Export intentionally disabled for now (admin-only + safe export later).
        exportEnabled: false,
      };
    },
    {
      errorMessage: "Failed to load audit logs",
      audit: false,
    },
    request,
  );
}
