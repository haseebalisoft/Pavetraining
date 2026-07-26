import { withAdminApi } from "@/lib/api/adminApi";
import { syncEventToOutlook } from "@/lib/services/calendar/calendarSyncService";

export const dynamic = "force-dynamic";
export const maxDuration = 60;

export async function POST(
  request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withAdminApi(
    "POST /api/admin/events/[id]/sync",
    async (admin, req) => {
      const result = await syncEventToOutlook(id, {
        force: true,
        userEmail: admin.loggedInEmail,
        request: req,
      });
      return { result };
    },
    {
      errorMessage: "Failed to sync event to Outlook",
      entityName: "Events",
      audit: false,
    },
    request,
  );
}
