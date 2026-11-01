import { withAdminApi } from "@/lib/api/adminApi";
import {
  deleteAdminRegister,
  updateAdminRegister,
} from "@/lib/services/adminCrudService";
import {
  triggerMatrixSyncAfterRegister,
  triggerMatrixSyncAfterRegisterDelete,
} from "@/lib/services/matrixSyncHook";
import {
  notifyTrainingRecordChange,
  readRegisterOutcome,
  summarizeTrainingManagerEmail,
} from "@/lib/services/trainingRecordNotificationService";

export const dynamic = "force-dynamic";

export async function PATCH(
  request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withAdminApi(
    "PATCH /api/admin/training-records/eusr/[id]",
    async (adminContext, req) => {
      const body = (await req.json()) as Record<string, unknown>;
      // Result before this edit: a re-save of the same Pass/Fail is not a new result.
      const previousOutcome = await readRegisterOutcome("eusrRegister", id);
      const { record, choiceWarnings } = await updateAdminRegister(
        "eusrRegister",
        id,
        body,
      );
      const matrixSync = await triggerMatrixSyncAfterRegister(
        "eusrRegister",
        record,
        adminContext.loggedInEmail,
      );
      const notification = await notifyTrainingRecordChange({
        register: "eusrRegister",
        action: "updated",
        record,
        actorEmail: adminContext.loggedInEmail,
        previousOutcome,
      });
      return {
        record,
        matrixSync, choiceWarnings,
        // Real delivery outcome ("sent" only when Graph accepted the email).
        trainingManagerEmail: summarizeTrainingManagerEmail(notification),
      };
    },
    { errorMessage: "Failed to update EUSR record" },
    request,
  );
}

export async function DELETE(
  _request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withAdminApi(
    "DELETE /api/admin/training-records/eusr/[id]",
    async (adminContext) => {
      const { deletedRecord } = await deleteAdminRegister("eusrRegister", id);
      const matrixSync = await triggerMatrixSyncAfterRegisterDelete(
        "eusrRegister",
        deletedRecord,
        adminContext.loggedInEmail,
      );
      return { ok: true, matrixSync, record: deletedRecord };
    },
    { errorMessage: "Failed to delete EUSR record" },
    _request,
  );
}
