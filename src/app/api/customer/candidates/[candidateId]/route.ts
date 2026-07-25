import { withCustomerApi } from "@/lib/api/customerApi";
import {
  AccessDeniedError,
  NotFoundError,
} from "@/lib/services/errorHandler";
import { assertCompanyMatch } from "@/lib/services/securityService";
import { getWorkforceById } from "@/lib/services/workforceService";

export const dynamic = "force-dynamic";

/**
 * Customer candidate profile.
 * Returns 403 when the candidate does not belong to the caller's company.
 */
export async function GET(
  request: Request,
  context: { params: Promise<{ candidateId: string }> },
) {
  const { candidateId } = await context.params;

  return withCustomerApi(
    "GET /api/customer/candidates/[candidateId]",
    async (customer) => {
      const candidate = await getWorkforceById(candidateId);
      if (!candidate) {
        throw new NotFoundError();
      }

      try {
        assertCompanyMatch(candidate.companyName, customer.companyName);
      } catch {
        throw new AccessDeniedError();
      }

      return {
        candidate: {
          id: candidate.id,
          candidateName: candidate.candidateName,
          companyName: candidate.companyName,
          workforceNumber: candidate.workforceNumber,
          department: candidate.department,
          // DOB is allowed but should stay de-emphasized in UI.
          dateOfBirth: candidate.dateOfBirth,
          status: candidate.status,
        },
      };
    },
    { entityName: "Workforce List", audit: true },
    request,
  );
}
