import { withAdminApi } from "@/lib/api/adminApi";
import {
  createAdminMatrix,
  listAdminMatrix,
} from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  const companyName = new URL(request.url).searchParams.get("companyName");
  return withAdminApi(
    "GET /api/admin/training-matrix",
    async () => ({ records: await listAdminMatrix(companyName) }),
    { errorMessage: "Failed to load training matrix" },
    request,
  );
}

export async function POST(request: Request) {
  return withAdminApi(
    "POST /api/admin/training-matrix",
    async (_context, req) => {
      const body = (await req.json()) as Record<string, unknown>;
      const record = await createAdminMatrix(body);
      return { record };
    },
    { errorMessage: "Failed to create matrix row" },
    request,
  );
}
