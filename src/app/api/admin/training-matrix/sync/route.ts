import { withAdminApi } from "@/lib/api/adminApi";
import {
  listAdminCompanies,
  listAdminWorkforce,
} from "@/lib/services/adminCrudService";
import {
  syncAllMatrix,
  syncCandidateMatrix,
  syncCompanyMatrix,
} from "@/lib/services/trainingMatrixSyncService";
import { ValidationError } from "@/lib/services/validationService";

export const dynamic = "force-dynamic";
export const maxDuration = 60;

function nameKey(value: string | null | undefined): string {
  return (value ?? "").trim().toLowerCase().replace(/\s+/g, " ");
}

export async function POST(request: Request) {
  return withAdminApi(
    "POST /api/admin/training-matrix/sync",
    async (context, req) => {
      const body = (await req.json().catch(() => ({}))) as {
        dryRun?: boolean;
        companyId?: string;
        candidateId?: string;
        candidateName?: string;
        companyName?: string;
      };
      const dryRun = Boolean(body.dryRun);
      const userEmail = context.loggedInEmail;

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
        });
        return { result };
      }

      const result = await syncAllMatrix({ dryRun, userEmail });
      return { result };
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
