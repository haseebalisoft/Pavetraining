import { withAdminApi } from "@/lib/api/adminApi";
import {
  createAdminRegister,
  listAdminRegister,
} from "@/lib/services/adminCrudService";

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
    async (_context, req) => {
      const body = (await req.json()) as Record<string, unknown>;
      const record = await createAdminRegister("inHouseCertificates", body);
      return { record };
    },
    { errorMessage: "Failed to create In-House record" },
    request,
  );
}
