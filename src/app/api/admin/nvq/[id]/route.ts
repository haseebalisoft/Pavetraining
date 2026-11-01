import { withAdminApi } from "@/lib/api/adminApi";
import {
  deleteAdminNvq,
  updateAdminNvq,
} from "@/lib/services/adminCrudService";
import { triggerMatrixSyncAfterNvq } from "@/lib/services/matrixSyncHook";
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
    "PATCH /api/admin/nvq/[id]",
    async (adminContext, req) => {
      const body = (await req.json()) as Record<string, unknown>;
      // Result before this edit: a re-save of the same Pass/Fail is not a new result.
      const previousOutcome = await readRegisterOutcome("nvqRegister", id);
      const record = await updateAdminNvq(id, body);
      const matrixSync = triggerMatrixSyncAfterNvq(record);
      const notification = await notifyTrainingRecordChange({
        register: "nvqRegister",
        action: "updated",
        record,
        actorEmail: adminContext.loggedInEmail,
        previousOutcome,
      });
      return {
        record,
        matrixSync,
        // Real delivery outcome ("sent" only when Graph accepted the email).
        trainingManagerEmail: summarizeTrainingManagerEmail(notification),
      };
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
