import { withAdminApi } from "@/lib/api/adminApi";
import { updateAdminRegister } from "@/lib/services/adminCrudService";
import { triggerMatrixSyncAfterRegister } from "@/lib/services/matrixSyncHook";

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
      const record = await updateAdminRegister("eusrRegister", id, body);
      const matrixSync = await triggerMatrixSyncAfterRegister(
        "eusrRegister",
        record,
        adminContext.loggedInEmail,
      );
      return { record, matrixSync };
    },
    { errorMessage: "Failed to update EUSR record" },
    request,
  );
}
