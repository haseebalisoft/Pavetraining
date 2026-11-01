import { withAdminApi } from "@/lib/api/adminApi";
import { createJobDeadline } from "@/lib/jobs/jobOutcome";
import { toAdminJobError } from "@/lib/services/errorHandler";
import { syncCompanyMatrix } from "@/lib/services/trainingMatrixSyncService";

export const dynamic = "force-dynamic";
/** Walks every candidate in the company serially. */
export const maxDuration = 300;

/** Stop starting new candidates 40s before maxDuration (see /sync route). */
const COMPANY_SYNC_BUDGET_MS = (300 - 40) * 1000;

export async function POST(
  request: Request,
  context: { params: Promise<{ companyId: string }> },
) {
  const deadline = createJobDeadline(COMPANY_SYNC_BUDGET_MS);
  const { companyId } = await context.params;
  return withAdminApi(
    "POST /api/admin/training-matrix/sync/company/[companyId]",
    async (adminContext, req) => {
      const body = (await req.json().catch(() => ({}))) as {
        dryRun?: boolean;
      };
      try {
        const result = await syncCompanyMatrix(companyId, {
          dryRun: Boolean(body.dryRun),
          userEmail: adminContext.loggedInEmail,
          deadline,
        });
        return { result };
      } catch (error) {
        throw toAdminJobError(error, "Company Training Matrix refresh");
      }
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
