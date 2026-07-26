import { withAdminApi } from "@/lib/api/adminApi";
import { updateAdminDocument } from "@/lib/services/adminCrudService";
import { triggerDocumentNotificationSafe } from "@/lib/services/documentNotificationService";

export const dynamic = "force-dynamic";

export async function PATCH(
  request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withAdminApi(
    "PATCH /api/admin/documents/[id]",
    async (adminContext, req) => {
      const body = (await req.json()) as Record<string, unknown>;
      const suppressNotifications = Boolean(body.suppressNotifications);
      const record = await updateAdminDocument(id, body);
      const notification = await triggerDocumentNotificationSafe(record, {
        suppressNotifications,
        actorEmail: adminContext.loggedInEmail,
      });
      return { record, notification };
    },
    { errorMessage: "Failed to update document" },
    request,
  );
}
