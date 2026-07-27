import { withAdminApi } from "@/lib/api/adminApi";
import { bulkDeleteAdminCompanies } from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

export async function POST(request: Request) {
  return withAdminApi(
    "POST /api/admin/companies/bulk-delete",
    async (_ctx, req) => {
      const body = (await req.json()) as { ids?: unknown };
      const ids = Array.isArray(body.ids)
        ? body.ids.map((id) => String(id))
        : [];
      const result = await bulkDeleteAdminCompanies(ids);
      return result;
    },
    { errorMessage: "Failed to bulk-delete companies" },
    request,
  );
}
