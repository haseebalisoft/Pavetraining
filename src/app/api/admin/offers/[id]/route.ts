import { withAdminApi } from "@/lib/api/adminApi";
import {
  deleteAdminOffer,
  updateAdminOffer,
} from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

export async function PATCH(
  request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withAdminApi(
    "PATCH /api/admin/offers/[id]",
    async (_ctx, req) => {
      const body = (await req.json()) as Record<string, unknown>;
      const record = await updateAdminOffer(id, body);
      return { record };
    },
    { errorMessage: "Failed to update offer" },
    request,
  );
}

export async function DELETE(
  _request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withAdminApi(
    "DELETE /api/admin/offers/[id]",
    async () => {
      await deleteAdminOffer(id);
      return { ok: true };
    },
    { errorMessage: "Failed to delete offer" },
    _request,
  );
}
