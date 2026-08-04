import { withAdminApi } from "@/lib/api/adminApi";
import {
  deleteAdminRegister,
  updateAdminRegister,
} from "@/lib/services/adminCrudService";
import { triggerMatrixSyncAfterRegister } from "@/lib/services/matrixSyncHook";

export const dynamic = "force-dynamic";

export async function PATCH(
  request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withAdminApi(
    "PATCH /api/admin/training-records/streetworks/[id]",
    async (adminContext, req) => {
      const body = (await req.json()) as Record<string, unknown>;
      const { record, choiceWarnings } = await updateAdminRegister(
        "nrswaRegister",
        id,
        body,
      );
      const matrixSync = await triggerMatrixSyncAfterRegister(
        "nrswaRegister",
        record,
        adminContext.loggedInEmail,
      );
      return { record, matrixSync, choiceWarnings };
    },
    { errorMessage: "Failed to update Streetworks record" },
    request,
  );
}

export async function DELETE(
  _request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withAdminApi(
    "DELETE /api/admin/training-records/streetworks/[id]",
    async () => {
      await deleteAdminRegister("nrswaRegister", id);
      return { ok: true };
    },
    { errorMessage: "Failed to delete Streetworks record" },
    _request,
  );
}
