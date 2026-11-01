import { withAdminApi } from "@/lib/api/adminApi";
import { createJobDeadline, type JobDeadline } from "@/lib/jobs/jobOutcome";
import {
  listAdminCompanies,
  listAdminWorkforce,
} from "@/lib/services/adminCrudService";
import { toAdminJobError } from "@/lib/services/errorHandler";
import {
  syncAllMatrix,
  syncCandidateMatrix,
  syncCompanyMatrix,
} from "@/lib/services/trainingMatrixSyncService";
import { ValidationError } from "@/lib/services/validationService";

export const dynamic = "force-dynamic";
/** A full refresh walks every candidate serially; it needs the long window. */
export const maxDuration = 300;

/**
 * Stop starting new candidates 40s before maxDuration so a response is always
 * sent (`result.incomplete` says how many were not reached) instead of the
 * platform killing the request with no answer.
 */
const MATRIX_SYNC_BUDGET_MS = (300 - 40) * 1000;

function nameKey(value: string | null | undefined): string {
  return (value ?? "").trim().toLowerCase().replace(/\s+/g, " ");
}

export async function POST(request: Request) {
  const deadline = createJobDeadline(MATRIX_SYNC_BUDGET_MS);
  return withAdminApi(
    "POST /api/admin/training-matrix/sync",
    async (context, req) => {
      try {
        return await runSync(context.loggedInEmail, req, deadline);
      } catch (error) {
        throw toAdminJobError(error, "Training Matrix refresh");
      }
    },
    {
      errorMessage: "Failed to sync Training Matrix",
      entityName: "training-matrix-sync",
      // trainingMatrixSyncService already writes MATRIX_SYNC_* audit rows —
      // avoid a duplicate generic ADMIN_CREATE entry per sync call.
      audit: false,
    },
    request,
  );
}

async function runSync(
  userEmail: string,
  req: Request,
  deadline: JobDeadline,
) {
  const body = (await req.json().catch(() => ({}))) as {
    dryRun?: boolean;
    companyId?: string;
    candidateId?: string;
    candidateName?: string;
    companyName?: string;
  };
  const dryRun = Boolean(body.dryRun);

  if (body.candidateId?.trim()) {
    const result = await syncCandidateMatrix(body.candidateId.trim(), {
      dryRun,
      userEmail,
    });
    return { result };
  }

  if (body.companyId?.trim()) {
    const result = await syncCompanyMatrix(body.companyId.trim(), {
      dryRun,
      userEmail,
      deadline,
    });
    return { result };
  }

  if (body.candidateName?.trim() && body.companyName?.trim()) {
    const [workforce, companies] = await Promise.all([
      listAdminWorkforce(),
      listAdminCompanies(),
    ]);
    const company =
      companies.find(
        (c) => nameKey(c.companyName) === nameKey(body.companyName),
      ) ?? null;
    const candidate =
      workforce.find(
        (w) =>
          nameKey(w.candidateName) === nameKey(body.candidateName) &&
          nameKey(w.companyName) === nameKey(body.companyName),
      ) ?? null;

    if (!candidate) {
      throw new ValidationError(
        "Could not resolve Workforce candidate for sync.",
      );
    }
    if (!company) {
      throw new ValidationError("Could not resolve company for sync.");
    }

    const result = await syncCandidateMatrix(candidate.id, {
      dryRun,
      userEmail,
    });
    return { result };
  }

  if (body.companyName?.trim() && !body.candidateName) {
    const companies = await listAdminCompanies();
    const company = companies.find(
      (c) => nameKey(c.companyName) === nameKey(body.companyName),
    );
    if (!company) {
      throw new ValidationError("Could not resolve company for sync.");
    }
    const result = await syncCompanyMatrix(company.id, {
      dryRun,
      userEmail,
      deadline,
    });
    return { result };
  }

  const result = await syncAllMatrix({ dryRun, userEmail, deadline });
  return { result };
}
