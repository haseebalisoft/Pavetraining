import { withAdminApi } from "@/lib/api/adminApi";
import {
  deleteAdminRegister,
  updateAdminRegister,
} from "@/lib/services/adminCrudService";
import {
  triggerMatrixSyncAfterRegister,
  triggerMatrixSyncAfterRegisterDelete,
} from "@/lib/services/matrixSyncHook";
import { notifyTrainingRecordChange } from "@/lib/services/trainingRecordNotificationService";

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
      await notifyTrainingRecordChange({
        register: "nrswaRegister",
        action: "updated",
        record,
        actorEmail: adminContext.loggedInEmail,
      });
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
    async (adminContext) => {
      const { deletedRecord } = await deleteAdminRegister("nrswaRegister", id);
      const matrixSync = await triggerMatrixSyncAfterRegisterDelete(
        "nrswaRegister",
        deletedRecord,
        adminContext.loggedInEmail,
      );
      return { ok: true, matrixSync, record: deletedRecord };
    },
    { errorMessage: "Failed to delete Streetworks record" },
    _request,
  );
}
