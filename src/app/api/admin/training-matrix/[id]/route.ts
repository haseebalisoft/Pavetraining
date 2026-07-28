import { withAdminApi } from "@/lib/api/adminApi";
import {
  deleteAdminMatrix,
  updateAdminMatrix,
} from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

export async function PATCH(
  request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withAdminApi(
    "PATCH /api/admin/training-matrix/[id]",
    async (_ctx, req) => {
      const body = (await req.json()) as Record<string, unknown>;
      const record = await updateAdminMatrix(id, body);
      return { record };
    },
    { errorMessage: "Failed to update matrix row" },
    request,
  );
}

export async function DELETE(
  request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withAdminApi(
    "DELETE /api/admin/training-matrix/[id]",
    async () => {
      await deleteAdminMatrix(id);
      return { deleted: true, id };
    },
    { errorMessage: "Failed to delete matrix row" },
    request,
  );
}
