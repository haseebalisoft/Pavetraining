import { withAdminApi, NotFoundError } from "@/lib/api/adminApi";
import { getAuditLogById } from "@/lib/services/auditLogService";

export const dynamic = "force-dynamic";

export async function GET(
  request: Request,
  context: { params: Promise<{ logId: string }> },
) {
  const { logId } = await context.params;
  return withAdminApi(
    "GET /api/admin/logs/[logId]",
    async () => {
      const log = await getAuditLogById(logId);
      if (!log) {
        throw new NotFoundError("Log entry not found.");
      }
      return { log };
    },
    {
      errorMessage: "Failed to load audit log",
      audit: false,
    },
    request,
  );
}
