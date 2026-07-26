import { withAdminApi, ValidationError } from "@/lib/api/adminApi";
import { bulkUpdateAdminDocuments } from "@/lib/services/adminCrudService";
import { logBulkUpload } from "@/lib/services/auditLogService";
import { triggerDocumentNotificationSafe } from "@/lib/services/documentNotificationService";
import { getSettings } from "@/lib/services/settingsService";

export const dynamic = "force-dynamic";

/**
 * Bulk-assign SharePoint metadata (Company, Candidate, Document Type,
 * Customer Visible, etc.) to selected Customer Documents items.
 * Defaults SuppressNotifications from Admin Settings.
 */
export async function PATCH(request: Request) {
  return withAdminApi(
    "PATCH /api/admin/documents/bulk",
    async (adminContext, req) => {
      const body = (await req.json()) as Record<string, unknown>;
      const idsRaw = body.ids;
      if (!Array.isArray(idsRaw)) {
        throw new ValidationError("ids must be an array of document IDs.");
      }
      const ids = idsRaw.map((id) => String(id));
      const suppressNotifications =
        body.suppressNotifications === undefined
          ? (await getSettings()).settings.suppressNotificationsDuringBulkUpload
          : Boolean(body.suppressNotifications);

      if (String(body.phase ?? "").toLowerCase() === "preview") {
        await logBulkUpload({
          userEmail: adminContext.loggedInEmail,
          phase: "preview",
          success: true,
          itemCount: ids.length,
          request: req,
        });
        return {
          preview: true,
          ids,
          suppressNotifications,
        };
      }

      const result = await bulkUpdateAdminDocuments(ids, body);

      await logBulkUpload({
        userEmail: adminContext.loggedInEmail,
        phase: "commit",
        success: result.failed.length === 0,
        itemCount: result.updated.length,
        errorMessage:
          result.failed.length > 0
            ? `${result.failed.length} document(s) failed to update`
            : null,
        request: req,
        metadata: { failedCount: result.failed.length },
      });

      if (!suppressNotifications) {
        for (const record of result.updated) {
          await triggerDocumentNotificationSafe(record, {
            suppressNotifications: false,
            actorEmail: adminContext.loggedInEmail,
          });
        }
      }

      return {
        ...result,
        suppressNotifications,
      };
    },
    {
      errorMessage: "Failed to update documents",
      // Dedicated bulk audit above — avoid duplicate ADMIN_UPDATE noise.
      audit: false,
      entityName: "Customer Documents",
    },
    request,
  );
}
