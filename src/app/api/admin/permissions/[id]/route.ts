import { withAdminApi } from "@/lib/api/adminApi";
import {
  deleteAdminPermission,
  updateAdminPermission,
} from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

export async function PATCH(
  request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withAdminApi(
    "PATCH /api/admin/permissions/[id]",
    async (_ctx, req) => {
      const body = (await req.json()) as Record<string, unknown>;
      const { record, choiceWarnings } = await updateAdminPermission(id, body);
      return { record, choiceWarnings };
    },
    { errorMessage: "Failed to update permission" },
    request,
  );
}

export async function DELETE(
  _request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withAdminApi(
    "DELETE /api/admin/permissions/[id]",
    async () => {
      await deleteAdminPermission(id);
      return { ok: true, id };
    },
    { errorMessage: "Failed to delete permission", entityName: "Permission" },
    _request,
  );
}
