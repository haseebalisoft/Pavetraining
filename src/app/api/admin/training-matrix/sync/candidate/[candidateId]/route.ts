import { withAdminApi } from "@/lib/api/adminApi";
import { syncCandidateMatrix } from "@/lib/services/trainingMatrixSyncService";

export const dynamic = "force-dynamic";
export const maxDuration = 60;

export async function POST(
  request: Request,
  context: { params: Promise<{ candidateId: string }> },
) {
  const { candidateId } = await context.params;
  return withAdminApi(
    "POST /api/admin/training-matrix/sync/candidate/[candidateId]",
    async (adminContext, req) => {
      const body = (await req.json().catch(() => ({}))) as {
        dryRun?: boolean;
      };
      const result = await syncCandidateMatrix(candidateId, {
        dryRun: Boolean(body.dryRun),
        userEmail: adminContext.loggedInEmail,
      });
      return { result };
    },
    {
      errorMessage: "Failed to sync candidate Training Matrix",
      entityName: "training-matrix-sync-candidate",
      // trainingMatrixSyncService already writes MATRIX_SYNC_* audit rows —
      // avoid a duplicate generic ADMIN_CREATE entry per sync call.
      audit: false,
    },
    request,
  );
}
