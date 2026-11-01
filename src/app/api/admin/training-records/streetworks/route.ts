import { withAdminApi } from "@/lib/api/adminApi";
import {
  createAdminRegister,
  listAdminRegister,
} from "@/lib/services/adminCrudService";
import { triggerMatrixSyncAfterRegister } from "@/lib/services/matrixSyncHook";
import {
  notifyTrainingRecordChange,
  summarizeTrainingManagerEmail,
} from "@/lib/services/trainingRecordNotificationService";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  const companyName = new URL(request.url).searchParams.get("companyName");
  return withAdminApi(
    "GET /api/admin/training-records/streetworks",
    async () => ({
      records: await listAdminRegister("nrswaRegister", companyName),
    }),
    { errorMessage: "Failed to load Streetworks records" },
    request,
  );
}

export async function POST(request: Request) {
  return withAdminApi(
    "POST /api/admin/training-records/streetworks",
    async (context, req) => {
      const body = (await req.json()) as Record<string, unknown>;
      const { record, choiceWarnings } = await createAdminRegister(
        "nrswaRegister",
        body,
      );
      const matrixSync = await triggerMatrixSyncAfterRegister(
        "nrswaRegister",
        record,
        context.loggedInEmail,
      );
      const notification = await notifyTrainingRecordChange({
        register: "nrswaRegister",
        action: "added",
        record,
        actorEmail: context.loggedInEmail,
      });
      return {
        record,
        matrixSync, choiceWarnings,
        // Real delivery outcome ("sent" only when Graph accepted the email).
        trainingManagerEmail: summarizeTrainingManagerEmail(notification),
      };
    },
    { errorMessage: "Failed to create Streetworks record" },
    request,
  );
}
