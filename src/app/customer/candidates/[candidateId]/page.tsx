import { redirect } from "next/navigation";

import { CandidateProfileView } from "@/components/customer/CandidateProfileView";
import { auth } from "@/auth";
import { assertCandidateAccess } from "@/lib/services/customerAccessService";
import { getCustomerContext } from "@/lib/services/customerContextService";
import { getCustomerMatrixRecords } from "@/lib/services/customerDashboardService";
import { assertCompanyMatch } from "@/lib/services/securityService";
import { getWorkforceById } from "@/lib/services/workforceService";
import type { CustomerContext } from "@/types/models";

export const dynamic = "force-dynamic";

export default async function CustomerCandidateProfilePage({
  params,
  searchParams,
}: {
  params: Promise<{ candidateId: string }>;
  searchParams: Promise<{ return?: string | string[] }>;
}) {
  const session = await auth();
  const email = session?.user?.email?.trim().toLowerCase();

  if (!email) {
    redirect("/login");
  }

  let context: CustomerContext;
  try {
    context = await getCustomerContext(email);
  } catch {
    redirect("/access-denied");
  }

  const { candidateId } = await params;
  const query = await searchParams;
  const returnRaw = Array.isArray(query.return) ? query.return[0] : query.return;
  const candidate = await getWorkforceById(candidateId);

  if (!candidate) {
    redirect("/customer/candidates");
  }

  try {
    assertCompanyMatch(candidate.companyName, context.companyName);
  } catch {
    redirect("/access-denied");
  }

  if (!assertCandidateAccess(candidate, context)) {
    redirect("/access-denied");
  }

  const matrixRows = await getCustomerMatrixRecords(
    context.companyName,
    context,
  );
  const matrixRow =
    matrixRows.find(
      (row) =>
        row.candidateName.trim().toLowerCase() ===
        candidate.candidateName.trim().toLowerCase(),
    ) ?? null;

  return (
    <CandidateProfileView
      candidate={candidate}
      matrixRow={matrixRow}
      matrixReturnHref={returnRaw}
    />
  );
}
