import { withAdminApi } from "@/lib/api/adminApi";
import { runExpiryReminderCheck } from "@/lib/services/expiryNotificationService";

export const dynamic = "force-dynamic";
export const maxDuration = 60;

export async function POST(request: Request) {
  return withAdminApi(
    "POST /api/admin/notifications/run-expiry-check",
    async (context, req) => {
      const body = (await req.json().catch(() => ({}))) as {
        dryRun?: boolean;
      };
      const result = await runExpiryReminderCheck({
        dryRun: Boolean(body.dryRun),
        actorEmail: context.loggedInEmail,
      });
      return { result };
    },
    {
      errorMessage: "Failed to run expiry reminder check",
      entityName: "notifications-expiry",
    },
    request,
  );
}
