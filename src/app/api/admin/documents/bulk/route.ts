import { withAdminApi, ValidationError } from "@/lib/api/adminApi";
import { bulkUpdateAdminDocuments } from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

/**
 * Bulk-assign SharePoint metadata (Company, Candidate, Document Type,
 * Customer Visible, etc.) to selected Customer Documents items.
 */
export async function PATCH(request: Request) {
  return withAdminApi(
    "PATCH /api/admin/documents/bulk",
    async (_ctx, req) => {
      const body = (await req.json()) as Record<string, unknown>;
      const idsRaw = body.ids;
      if (!Array.isArray(idsRaw)) {
        throw new ValidationError("ids must be an array of document IDs.");
      }
      const ids = idsRaw.map((id) => String(id));
      const result = await bulkUpdateAdminDocuments(ids, body);
      return result;
    },
    {
      errorMessage: "Failed to update documents",
      audit: true,
      entityName: "Customer Documents",
    },
    request,
  );
}
