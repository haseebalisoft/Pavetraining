import { withSharePointAdminApi } from "@/lib/api/adminApi";
import {
  deleteAdminPermission,
  updateAdminPermission,
} from "@/lib/services/adminCrudService";
import { logPermissionDepartmentScopeUpdate } from "@/lib/services/auditLogService";

export const dynamic = "force-dynamic";

export async function PATCH(
  request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withSharePointAdminApi(
    "PATCH /api/admin/permissions/[id]",
    async (ctx, req) => {
      const body = (await req.json()) as Record<string, unknown>;
      const { record, choiceWarnings } = await updateAdminPermission(id, body);
      if (body.departmentsAllowed !== undefined) {
        await logPermissionDepartmentScopeUpdate({
          userEmail: ctx.loggedInEmail,
          permissionId: record.id,
          personName: record.name,
          departmentNames: record.departmentScopes,
          companyName: record.companyName,
          success: true,
          request: req,
        });
      }
      return { record, choiceWarnings };
    },
    { errorMessage: "Failed to update permission" },
    request,
  );
}

export async function DELETE(
  _request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withSharePointAdminApi(
    "DELETE /api/admin/permissions/[id]",
    async () => {
      const record = await deleteAdminPermission(id);
      return { ok: true, id, record };
    },
    { errorMessage: "Failed to delete permission", entityName: "Permission" },
    _request,
  );
}
