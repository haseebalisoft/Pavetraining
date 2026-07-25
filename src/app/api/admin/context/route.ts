import { withAdminApi } from "@/lib/api/adminApi";

export const dynamic = "force-dynamic";

/**
 * Admin context from Permissions List.
 * Only RoleType = Admin may access. Non-admins receive 403.
 */
export async function GET() {
  return withAdminApi("GET /api/admin/context", async (context) => context, {
    errorMessage: "Failed to resolve admin context",
  });
}
