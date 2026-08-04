import { withAdminApi } from "@/lib/api/adminApi";
import {
  deleteAdminNvq,
  updateAdminNvq,
} from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

export async function PATCH(
  request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withAdminApi(
    "PATCH /api/admin/nvq/[id]",
    async (_ctx, req) => {
      const body = (await req.json()) as Record<string, unknown>;
      const record = await updateAdminNvq(id, body);
      return { record };
    },
    { errorMessage: "Failed to update NVQ record" },
    request,
  );
}

export async function DELETE(
  _request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withAdminApi(
    "DELETE /api/admin/nvq/[id]",
    async () => {
      await deleteAdminNvq(id);
      return { ok: true };
    },
    { errorMessage: "Failed to delete NVQ record" },
    _request,
  );
}
