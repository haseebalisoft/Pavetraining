import { withAdminApi } from "@/lib/api/adminApi";
import {
  deleteAdminDepartment,
  updateAdminDepartment,
} from "@/lib/services/departmentService";

export const dynamic = "force-dynamic";

export async function PATCH(
  request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withAdminApi(
    "PATCH /api/admin/departments/[id]",
    async (_ctx, req) => {
      const body = (await req.json()) as Record<string, unknown>;
      const record = await updateAdminDepartment(id, {
        name: body.name == null ? undefined : String(body.name),
        companyId:
          body.companyId == null ? undefined : String(body.companyId),
      });
      return { record };
    },
    { errorMessage: "Failed to update department" },
    request,
  );
}

export async function DELETE(
  _request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withAdminApi(
    "DELETE /api/admin/departments/[id]",
    async () => {
      await deleteAdminDepartment(id);
      return { ok: true };
    },
    { errorMessage: "Failed to delete department" },
    _request,
  );
}
