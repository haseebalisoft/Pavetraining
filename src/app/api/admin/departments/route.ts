import { withAdminApi } from "@/lib/api/adminApi";
import {
  createAdminDepartment,
  listAdminDepartments,
} from "@/lib/services/departmentService";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  return withAdminApi(
    "GET /api/admin/departments",
    async (_ctx, req) => {
      const url = new URL(req.url);
      const companyId = url.searchParams.get("companyId");
      return { records: await listAdminDepartments(companyId) };
    },
    { errorMessage: "Failed to load departments" },
    request,
  );
}

export async function POST(request: Request) {
  return withAdminApi(
    "POST /api/admin/departments",
    async (_ctx, req) => {
      const body = (await req.json()) as Record<string, unknown>;
      const record = await createAdminDepartment({
        name: String(body.name ?? ""),
        companyId: String(body.companyId ?? body.companyName ?? ""),
      });
      return { record };
    },
    { errorMessage: "Failed to create department" },
    request,
  );
}
