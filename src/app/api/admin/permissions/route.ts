import { withAdminApi } from "@/lib/api/adminApi";
import {
  createAdminPermission,
  listAdminPermissions,
} from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  return withAdminApi(
    "GET /api/admin/permissions",
    async () => ({ records: await listAdminPermissions() }),
    { errorMessage: "Failed to load permissions" },
    request,
  );
}

export async function POST(request: Request) {
  return withAdminApi(
    "POST /api/admin/permissions",
    async (_context, req) => {
      const body = (await req.json()) as Record<string, unknown>;
      const record = await createAdminPermission(body);
      return { record };
    },
    { errorMessage: "Failed to create permission" },
    request,
  );
}
