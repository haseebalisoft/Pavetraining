import { withAdminApi } from "@/lib/api/adminApi";
import {
  createAdminRegister,
  listAdminRegister,
} from "@/lib/services/adminCrudService";
import { triggerMatrixSyncAfterRegister } from "@/lib/services/matrixSyncHook";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  const companyName = new URL(request.url).searchParams.get("companyName");
  return withAdminApi(
    "GET /api/admin/training-records/npors",
    async () => ({
      records: await listAdminRegister("nporsRegister", companyName),
    }),
    { errorMessage: "Failed to load NPORS records" },
    request,
  );
}

export async function POST(request: Request) {
  return withAdminApi(
    "POST /api/admin/training-records/npors",
    async (context, req) => {
      const body = (await req.json()) as Record<string, unknown>;
      const record = await createAdminRegister("nporsRegister", body);
      const matrixSync = await triggerMatrixSyncAfterRegister(
        "nporsRegister",
        record,
        context.loggedInEmail,
      );
      return { record, matrixSync };
    },
    { errorMessage: "Failed to create NPORS record" },
    request,
  );
}
