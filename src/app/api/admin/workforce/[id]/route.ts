import { withAdminApi } from "@/lib/api/adminApi";
import {
  deleteAdminWorkforce,
  updateAdminWorkforce,
} from "@/lib/services/adminCrudService";
import { logWorkforceDepartmentAssign } from "@/lib/services/auditLogService";

export const dynamic = "force-dynamic";

export async function PATCH(
  request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withAdminApi(
    "PATCH /api/admin/workforce/[id]",
    async (ctx, req) => {
      const body = (await req.json()) as Record<string, unknown>;
      const record = await updateAdminWorkforce(id, body);
      if (body.department !== undefined || body.departmentText !== undefined) {
        await logWorkforceDepartmentAssign({
          userEmail: ctx.loggedInEmail,
          workforceId: record.id,
          candidateName: record.candidateName,
          departmentName: record.department ?? "(none)",
          companyName: record.companyName,
          success: true,
          request: req,
        });
      }
      return { record };
    },
    { errorMessage: "Failed to update candidate" },
    request,
  );
}

export async function DELETE(
  _request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withAdminApi(
    "DELETE /api/admin/workforce/[id]",
    async () => {
      const record = await deleteAdminWorkforce(id);
      return { ok: true, record };
    },
    { errorMessage: "Failed to delete candidate" },
    _request,
  );
}
