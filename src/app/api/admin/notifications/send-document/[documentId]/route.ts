import { withAdminApi } from "@/lib/api/adminApi";
import { notifyDocumentById } from "@/lib/services/documentNotificationService";

export const dynamic = "force-dynamic";
export const maxDuration = 60;

export async function POST(
  request: Request,
  context: { params: Promise<{ documentId: string }> },
) {
  const { documentId } = await context.params;
  return withAdminApi(
    "POST /api/admin/notifications/send-document/[documentId]",
    async (adminContext, req) => {
      const body = (await req.json().catch(() => ({}))) as {
        force?: boolean;
        suppressNotifications?: boolean;
      };
      const result = await notifyDocumentById(documentId, {
        force: Boolean(body.force),
        suppressNotifications: Boolean(body.suppressNotifications),
        actorEmail: adminContext.loggedInEmail,
      });
      return { result };
    },
    {
      errorMessage: "Failed to send document notification",
      entityName: "notifications-document",
    },
    request,
  );
}
