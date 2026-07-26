import { withAdminApi, ValidationError } from "@/lib/api/adminApi";
import { getNotificationSettings } from "@/lib/services/notificationConfig";
import { listNotificationLogs } from "@/lib/services/notificationLogService";
import { sendTestNotification } from "@/lib/services/notificationService";

export const dynamic = "force-dynamic";

export async function POST(request: Request) {
  return withAdminApi(
    "POST /api/admin/notifications/test",
    async (context, req) => {
      const body = (await req.json().catch(() => ({}))) as {
        to?: string;
      };
      const to =
        body.to?.trim() ||
        context.loggedInEmail ||
        (await getNotificationSettings()).fromEmail;
      if (!to) {
        throw new ValidationError(
          "Provide a recipient email, or configure NOTIFICATION_FROM_EMAIL.",
        );
      }

      const result = await sendTestNotification({
        to,
        actorEmail: context.loggedInEmail,
      });
      return { result, settings: await getNotificationSettings() };
    },
    {
      errorMessage: "Failed to send test notification",
      entityName: "notifications-test",
    },
    request,
  );
}
