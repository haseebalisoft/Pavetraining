import { withAdminApi } from "@/lib/api/adminApi";
import { listAuditLogs } from "@/lib/services/auditLogService";
import type { AuditLogQuery } from "@/types/audit";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  const url = new URL(request.url);
  const query: AuditLogQuery = {
    search: url.searchParams.get("search"),
    action: url.searchParams.get("action"),
    success: (url.searchParams.get("success") as AuditLogQuery["success"]) ?? "all",
    entityType: url.searchParams.get("entityType"),
    from: url.searchParams.get("from"),
    to: url.searchParams.get("to"),
    top: Number(url.searchParams.get("top") ?? "200"),
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
