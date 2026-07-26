import { withAdminApi } from "@/lib/api/adminApi";
import { syncEventToOutlook } from "@/lib/services/calendar/calendarSyncService";
import { writeAuditEvent } from "@/lib/services/auditLogService";

export const dynamic = "force-dynamic";
export const maxDuration = 60;

export async function POST(
  request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withAdminApi(
    "POST /api/admin/events/[id]/retry-sync",
    async (admin, req) => {
      await writeAuditEvent({
        userEmail: admin.loggedInEmail,
        roleType: "Admin",
        action: "EVENT_SYNC_RETRY",
        entityType: "Events",
        entityId: id,
        success: true,
        request: req,
      });

      const result = await syncEventToOutlook(id, {
        force: true,
        userEmail: admin.loggedInEmail,
        request: req,
      });
      return { result };
    },
    {
      errorMessage: "Failed to retry Outlook sync",
      entityName: "Events",
      audit: false,
    },
    request,
  );
}
