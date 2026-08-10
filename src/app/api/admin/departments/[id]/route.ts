import { withAdminApi } from "@/lib/api/adminApi";
import { logDepartmentChange } from "@/lib/services/auditLogService";
import {
  deleteAdminDepartment,
  normalizeDepartmentStatus,
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
    async (ctx, req) => {
      const body = (await req.json()) as Record<string, unknown>;
      try {
        const { record, previousStatus } = await updateAdminDepartment(id, {
          name: body.name == null ? undefined : String(body.name),
          companyId:
            body.companyId == null ? undefined : String(body.companyId),
          status:
            body.status === undefined
              ? undefined
              : normalizeDepartmentStatus(String(body.status)),
          notes:
            body.notes === undefined
              ? undefined
              : body.notes === null
                ? null
                : String(body.notes),
        });
        const deactivated =
          previousStatus !== "Inactive" && record.status === "Inactive";
        await logDepartmentChange({
          action: deactivated ? "DEPARTMENT_DEACTIVATE" : "DEPARTMENT_UPDATE",
          userEmail: ctx.loggedInEmail,
          departmentId: record.id,
          departmentName: record.name,
          companyName: record.companyName,
          success: true,
          request: req,
        });
        return { record };
      } catch (error) {
        await logDepartmentChange({
          action: "DEPARTMENT_UPDATE",
          userEmail: ctx.loggedInEmail,
          departmentId: id,
          departmentName: body.name == null ? "" : String(body.name),
          success: false,
          errorMessage: error instanceof Error ? error.message : "Update failed",
          request: req,
        });
        throw error;
      }
    },
    { errorMessage: "Failed to update department", audit: false },
    request,
  );
}

export async function DELETE(
  request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withAdminApi(
    "DELETE /api/admin/departments/[id]",
    async (ctx, req) => {
      try {
        const deleted = await deleteAdminDepartment(id);
        await logDepartmentChange({
          action: "DEPARTMENT_DELETE",
          userEmail: ctx.loggedInEmail,
          departmentId: id,
          departmentName: deleted.name,
          companyName: deleted.companyName,
          success: true,
          request: req,
        });
        return { ok: true };
      } catch (error) {
        await logDepartmentChange({
          action: "DEPARTMENT_DELETE",
          userEmail: ctx.loggedInEmail,
          departmentId: id,
          departmentName: "",
          success: false,
          errorMessage: error instanceof Error ? error.message : "Delete failed",
          request: req,
        });
        throw error;
      }
    },
    { errorMessage: "Failed to delete department", audit: false },
    request,
  );
}
