import { withAdminApi } from "@/lib/api/adminApi";
import { logDepartmentChange } from "@/lib/services/auditLogService";
import {
  createAdminDepartment,
  listAdminDepartments,
  normalizeDepartmentStatus,
} from "@/lib/services/departmentService";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  return withAdminApi(
    "GET /api/admin/departments",
    async (_ctx, req) => {
      const url = new URL(req.url);
      const companyId = url.searchParams.get("companyId");
      // Backs the admin Departments table — includes Inactive so they stay
      // visible and reactivatable. Pickers use their own Active-only reads.
      return {
        records: await listAdminDepartments(companyId, {
          includeInactive: true,
        }),
      };
    },
    { errorMessage: "Failed to load departments" },
    request,
  );
}

export async function POST(request: Request) {
  return withAdminApi(
    "POST /api/admin/departments",
    async (context, req) => {
      const body = (await req.json()) as Record<string, unknown>;
      const name = String(body.name ?? "");
      try {
        const record = await createAdminDepartment({
          name,
          companyId: String(body.companyId ?? body.companyName ?? ""),
          status:
            body.status === undefined
              ? undefined
              : normalizeDepartmentStatus(String(body.status)),
          notes:
            body.notes === null || body.notes === undefined
              ? null
              : String(body.notes),
        });
        await logDepartmentChange({
          action: "DEPARTMENT_CREATE",
          userEmail: context.loggedInEmail,
          departmentId: record.id,
          departmentName: record.name,
          companyName: record.companyName,
          success: true,
          request: req,
        });
        return { record };
      } catch (error) {
        await logDepartmentChange({
          action: "DEPARTMENT_CREATE",
          userEmail: context.loggedInEmail,
          departmentName: name,
          success: false,
          errorMessage: error instanceof Error ? error.message : "Create failed",
          request: req,
        });
        throw error;
      }
    },
    { errorMessage: "Failed to create department", audit: false },
    request,
  );
}
