import { withAdminApi } from "@/lib/api/adminApi";
import {
  createAdminRegister,
  listAdminRegister,
} from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  const companyName = new URL(request.url).searchParams.get("companyName");
  return withAdminApi(
    "GET /api/admin/training-records/eusr",
    async () => ({
      records: await listAdminRegister("eusrRegister", companyName),
    }),
    { errorMessage: "Failed to load EUSR records" },
    request,
  );
}

export async function POST(request: Request) {
  return withAdminApi(
    "POST /api/admin/training-records/eusr",
    async (_context, req) => {
      const body = (await req.json()) as Record<string, unknown>;
      const record = await createAdminRegister("eusrRegister", body);
      return { record };
    },
    { errorMessage: "Failed to create EUSR record" },
    request,
  );
}
