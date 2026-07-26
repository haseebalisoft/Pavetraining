import { withAdminApi } from "@/lib/api/adminApi";
import { syncPendingEvents } from "@/lib/services/calendar/calendarSyncService";

export const dynamic = "force-dynamic";
export const maxDuration = 60;

export async function POST(request: Request) {
  return withAdminApi(
    "POST /api/admin/events/sync-pending",
    async (admin, req) => {
      const body = (await req.json().catch(() => ({}))) as { limit?: number };
      const result = await syncPendingEvents({
        userEmail: admin.loggedInEmail,
        request: req,
        limit: typeof body.limit === "number" ? body.limit : undefined,
      });
      return result;
    },
    {
      errorMessage: "Failed to sync pending events",
      entityName: "Events",
    },
    request,
  );
}
