import { withAdminApi } from "@/lib/api/adminApi";
import { syncCompanyMatrix } from "@/lib/services/trainingMatrixSyncService";

export const dynamic = "force-dynamic";
export const maxDuration = 60;

export async function POST(
  request: Request,
  context: { params: Promise<{ companyId: string }> },
) {
  const { companyId } = await context.params;
  return withAdminApi(
    "POST /api/admin/training-matrix/sync/company/[companyId]",
    async (adminContext, req) => {
      const body = (await req.json().catch(() => ({}))) as {
        dryRun?: boolean;
      };
      const result = await syncCompanyMatrix(companyId, {
        dryRun: Boolean(body.dryRun),
        userEmail: adminContext.loggedInEmail,
      });
      return { result };
    },
    {
      errorMessage: "Failed to sync company Training Matrix",
      entityName: "training-matrix-sync-company",
      // trainingMatrixSyncService already writes MATRIX_SYNC_* audit rows —
      // avoid a duplicate generic ADMIN_CREATE entry per sync call.
      audit: false,
    },
    request,
  );
}
