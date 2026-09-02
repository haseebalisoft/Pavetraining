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
    "PATCH /api/admin/training-records/in-house/[id]",
    async (adminContext, req) => {
      const body = (await req.json()) as Record<string, unknown>;
      const { record, choiceWarnings } = await updateAdminRegister(
        "inHouseCertificates",
        id,
        body,
      );
      const matrixSync = await triggerMatrixSyncAfterRegister(
        "inHouseCertificates",
        record,
        adminContext.loggedInEmail,
      );
      await notifyTrainingRecordChange({
        register: "inHouseCertificates",
        action: "updated",
        record,
        actorEmail: adminContext.loggedInEmail,
      });
      return { record, matrixSync, choiceWarnings };
    },
    { errorMessage: "Failed to update In-House record" },
    request,
  );
}

export async function DELETE(
  _request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withAdminApi(
    "DELETE /api/admin/training-records/in-house/[id]",
    async (adminContext) => {
      const { deletedRecord } = await deleteAdminRegister(
        "inHouseCertificates",
        id,
      );
      const matrixSync = await triggerMatrixSyncAfterRegisterDelete(
        "inHouseCertificates",
        deletedRecord,
        adminContext.loggedInEmail,
      );
      return { ok: true, matrixSync, record: deletedRecord };
    },
    { errorMessage: "Failed to delete In-House record" },
    _request,
  );
}
