import { withAdminApi } from "@/lib/api/adminApi";
import {
  deleteAdminWorkforce,
  updateAdminWorkforce,
} from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

export async function PATCH(
  request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withAdminApi(
    "PATCH /api/admin/workforce/[id]",
    async (_ctx, req) => {
      const body = (await req.json()) as Record<string, unknown>;
      const record = await updateAdminWorkforce(id, body);
      return { record };
    },
    { errorMessage: "Failed to update candidate" },
    request,
  );
}

export async function DELETE(
  _request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withAdminApi(
    "DELETE /api/admin/workforce/[id]",
    async () => {
      await deleteAdminWorkforce(id);
      return { ok: true };
    },
    { errorMessage: "Failed to delete candidate" },
    _request,
  );
}
