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
    "PATCH /api/admin/training-records/npors/[id]",
    async (adminContext, req) => {
      const body = (await req.json()) as Record<string, unknown>;
      const { record, choiceWarnings } = await updateAdminRegister(
        "nporsRegister",
        id,
        body,
      );
      const matrixSync = await triggerMatrixSyncAfterRegister(
        "nporsRegister",
        record,
        adminContext.loggedInEmail,
      );
      // Best-effort — never blocks the save.
      await notifyTrainingRecordChange({
        register: "nporsRegister",
        action: "updated",
        record,
        actorEmail: adminContext.loggedInEmail,
      });
      return { record, matrixSync, choiceWarnings };
    },
    { errorMessage: "Failed to update NPORS record" },
    request,
  );
}

export async function DELETE(
  _request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withAdminApi(
    "DELETE /api/admin/training-records/npors/[id]",
    async (adminContext) => {
      const { deletedRecord } = await deleteAdminRegister("nporsRegister", id);
      const matrixSync = await triggerMatrixSyncAfterRegisterDelete(
        "nporsRegister",
        deletedRecord,
        adminContext.loggedInEmail,
      );
      return { ok: true, matrixSync, record: deletedRecord };
    },
    { errorMessage: "Failed to delete NPORS record" },
    _request,
  );
}
