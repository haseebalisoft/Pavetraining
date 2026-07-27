import { withAdminApi } from "@/lib/api/adminApi";
import {
  deleteAdminCompany,
  updateAdminCompany,
} from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

export async function PATCH(
  request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withAdminApi(
    "PATCH /api/admin/companies/[id]",
    async (_ctx, req) => {
      const body = (await req.json()) as Record<string, unknown>;
      const company = await updateAdminCompany(id, body);
      return { company };
    },
    { errorMessage: "Failed to update company" },
    request,
  );
}

export async function DELETE(
  request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withAdminApi(
    "DELETE /api/admin/companies/[id]",
    async () => {
      const result = await deleteAdminCompany(id);
      return { result };
    },
    { errorMessage: "Failed to delete company" },
    request,
  );
}
