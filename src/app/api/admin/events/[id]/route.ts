import { withAdminApi } from "@/lib/api/adminApi";
import { updateAdminEvent } from "@/lib/services/adminCrudService";

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
      const record = await updateAdminEvent(id, body);
      return { record };
    },
    { errorMessage: "Failed to update event" },
    request,
  );
}
