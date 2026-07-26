import { withAdminApi } from "@/lib/api/adminApi";
import { getEventSyncStatusSummary } from "@/lib/services/calendar/calendarSyncService";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  return withAdminApi(
    "GET /api/admin/events/sync-status",
    async () => {
      const summary = await getEventSyncStatusSummary();
      return { summary };
    },
    { errorMessage: "Failed to load event sync status" },
    request,
  );
}
