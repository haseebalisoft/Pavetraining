import { withAdminApi } from "@/lib/api/adminApi";
import {
  createAdminWorkforce,
  listAdminWorkforce,
} from "@/lib/services/adminCrudService";
import { logWorkforceDepartmentAssign } from "@/lib/services/auditLogService";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  const companyName = new URL(request.url).searchParams.get("companyName");
  return withAdminApi(
    "GET /api/admin/workforce",
    async () => ({
      records: await listAdminWorkforce(companyName),
    }),
    { errorMessage: "Failed to load workforce" },
    request,
  );
}

export async function POST(request: Request) {
  return withAdminApi(
    "POST /api/admin/workforce",
    async (context, req) => {
      const body = (await req.json()) as Record<string, unknown>;
      const record = await createAdminWorkforce(body);
      if (body.department !== undefined || body.departmentText !== undefined) {
        await logWorkforceDepartmentAssign({
          userEmail: context.loggedInEmail,
          workforceId: record.id,
          candidateName: record.candidateName,
          departmentName: record.department ?? "(none)",
          companyName: record.companyName,
          success: true,
          request: req,
        });
      }
      const folderWarning =
        "folderWarning" in record
          ? (record as { folderWarning?: string }).folderWarning
          : undefined;
      const matrixSeedWarning =
        "matrixSeedWarning" in record
          ? (record as { matrixSeedWarning?: string }).matrixSeedWarning
          : undefined;
      const warnings = [folderWarning, matrixSeedWarning]
        .map((part) => part?.trim())
        .filter(Boolean);
      return {
        record,
        warning: warnings.length ? warnings.join(" ") : undefined,
        matrixSeedWarning,
      };
    },
    { errorMessage: "Failed to create candidate" },
    request,
  );
}
