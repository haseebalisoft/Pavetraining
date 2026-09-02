import { withAdminApi } from "@/lib/api/adminApi";
import {
  deleteAdminNvq,
  updateAdminNvq,
} from "@/lib/services/adminCrudService";
import { triggerMatrixSyncAfterNvq } from "@/lib/services/matrixSyncHook";
import { notifyTrainingRecordChange } from "@/lib/services/trainingRecordNotificationService";

export const dynamic = "force-dynamic";

export async function PATCH(
  request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withAdminApi(
    "PATCH /api/admin/nvq/[id]",
    async (adminContext, req) => {
      const body = (await req.json()) as Record<string, unknown>;
      const record = await updateAdminNvq(id, body);
      const matrixSync = triggerMatrixSyncAfterNvq(record);
      await notifyTrainingRecordChange({
        register: "nvqRegister",
        action: "updated",
        record,
        actorEmail: adminContext.loggedInEmail,
      });
      return { record, matrixSync };
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
