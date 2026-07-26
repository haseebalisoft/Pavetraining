import { withAdminApi } from "@/lib/api/adminApi";
import {
  createAdminCompany,
  listAdminCompanies,
} from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  return withAdminApi(
    "GET /api/admin/companies",
    async () => ({ companies: await listAdminCompanies() }),
    { errorMessage: "Failed to load companies" },
    request,
  );
}

export async function POST(request: Request) {
  return withAdminApi(
    "POST /api/admin/companies",
    async (_context, req) => {
      const body = (await req.json()) as Record<string, unknown>;
      const company = await createAdminCompany(body);
      return {
        company,
        warning:
          "folderWarning" in company
            ? (company as { folderWarning?: string }).folderWarning
            : undefined,
      };
    },
    { errorMessage: "Failed to create company" },
    request,
  );
}
