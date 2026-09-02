import { withAdminApi } from "@/lib/api/adminApi";
import {
  createAdminRegister,
  listAdminRegister,
} from "@/lib/services/adminCrudService";
import { triggerMatrixSyncAfterRegister } from "@/lib/services/matrixSyncHook";
import { notifyTrainingRecordChange } from "@/lib/services/trainingRecordNotificationService";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  const companyName = new URL(request.url).searchParams.get("companyName");
  return withAdminApi(
    "GET /api/admin/training-records/in-house",
    async () => ({
      records: await listAdminRegister("inHouseCertificates", companyName),
    }),
    { errorMessage: "Failed to load In-House records" },
    request,
  );
}

export async function POST(request: Request) {
  return withAdminApi(
    "POST /api/admin/training-records/in-house",
    async (context, req) => {
      const body = (await req.json()) as Record<string, unknown>;
      const { record, choiceWarnings } = await createAdminRegister(
        "inHouseCertificates",
        body,
      );
      // Asbestos Awareness Pass syncs to N031; other courses stay standalone.
      const matrixSync = await triggerMatrixSyncAfterRegister(
        "inHouseCertificates",
        record,
        context.loggedInEmail,
      );
      await notifyTrainingRecordChange({
        register: "inHouseCertificates",
        action: "added",
        record,
        actorEmail: context.loggedInEmail,
      });
      return { record, matrixSync, choiceWarnings };
    },
    { errorMessage: "Failed to create In-House record" },
    request,
  );
}
