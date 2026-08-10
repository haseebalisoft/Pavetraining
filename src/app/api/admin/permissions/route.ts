import { withAdminApi } from "@/lib/api/adminApi";
import {
  createAdminPermission,
  listAdminPermissions,
} from "@/lib/services/adminCrudService";
import { logPermissionDepartmentScopeUpdate } from "@/lib/services/auditLogService";

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
    async (context, req) => {
      const body = (await req.json()) as Record<string, unknown>;
      const record = await createAdminPermission(body);
      if (body.departmentsAllowed !== undefined) {
        await logPermissionDepartmentScopeUpdate({
          userEmail: context.loggedInEmail,
          permissionId: record.id,
          personName: record.name,
          departmentNames: record.departmentScopes,
          companyName: record.companyName,
          success: true,
          request: req,
        });
      }
      return { record };
    },
    { errorMessage: "Failed to create permission" },
    request,
  );
}
