import { withAdminApi } from "@/lib/api/adminApi";
import { createAdminNvq, listAdminNvq } from "@/lib/services/adminCrudService";
import { triggerMatrixSyncAfterNvq } from "@/lib/services/matrixSyncHook";
import { notifyTrainingRecordChange } from "@/lib/services/trainingRecordNotificationService";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  const companyName = new URL(request.url).searchParams.get("companyName");
  return withAdminApi(
    "GET /api/admin/nvq",
    async () => ({ records: await listAdminNvq(companyName) }),
    { errorMessage: "Failed to load NVQ records" },
    request,
  );
}

export async function POST(request: Request) {
  return withAdminApi(
    "POST /api/admin/nvq",
    async (context, req) => {
      const body = (await req.json()) as Record<string, unknown>;
      const record = await createAdminNvq(body);
      const matrixSync = triggerMatrixSyncAfterNvq(record);
      await notifyTrainingRecordChange({
        register: "nvqRegister",
        action: "added",
        record,
        actorEmail: context.loggedInEmail,
      });
      return { record, matrixSync };
    },
    { errorMessage: "Failed to create NVQ record" },
    request,
  );
}
