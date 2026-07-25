import { withAdminApi } from "@/lib/api/adminApi";
import {
  createAdminWorkforce,
  listAdminWorkforce,
} from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  const companyName = new URL(request.url).searchParams.get("companyName");
  return withAdminApi(
    "GET /api/admin/workforce",
    async () => ({
      records: await listAdminWorkforce(companyName),
    }),
    { errorMessage: "Failed to load workforce" },
    request,
  );
}

export async function POST(request: Request) {
  return withAdminApi(
    "POST /api/admin/workforce",
    async (_context, req) => {
      const body = (await req.json()) as Record<string, unknown>;
      const record = await createAdminWorkforce(body);
      return { record };
    },
    { errorMessage: "Failed to create candidate" },
    request,
  );
}
