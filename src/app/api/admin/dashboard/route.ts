import { withAdminApi } from "@/lib/api/adminApi";
import { getAdminDashboard } from "@/lib/services/adminDashboardService";

export const dynamic = "force-dynamic";

/**
 * Admin dashboard metrics from SharePoint.
 * Optional ?companyId= filters operational stats to one company.
 * Only RoleType = Admin may access. Non-admins receive 403.
 */
export async function GET(request: Request) {
  const { searchParams } = new URL(request.url);
  const companyId = searchParams.get("companyId");

  return withAdminApi(
    "GET /api/admin/dashboard",
    async () => getAdminDashboard(companyId),
    { errorMessage: "Failed to load admin dashboard" },
  );
}
