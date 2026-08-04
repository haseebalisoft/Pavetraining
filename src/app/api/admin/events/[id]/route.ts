import { withAdminApi } from "@/lib/api/adminApi";
import {
  deleteAdminEvent,
  updateAdminEvent,
} from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

export async function PATCH(
  request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withAdminApi(
    "PATCH /api/admin/events/[id]",
    async (_ctx, req) => {
      const body = (await req.json()) as Record<string, unknown>;
      return updateAdminEvent(id, body);
    },
    { errorMessage: "Failed to update event" },
    request,
  );
}

export async function DELETE(
  _request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withAdminApi(
    "DELETE /api/admin/events/[id]",
    async () => {
      await deleteAdminEvent(id);
      return { ok: true };
    },
    { errorMessage: "Failed to delete event" },
    _request,
  );
}
