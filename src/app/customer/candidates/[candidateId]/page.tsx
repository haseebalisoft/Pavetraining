import { redirect } from "next/navigation";

import { CandidateProfileView } from "@/components/customer/CandidateProfileView";
import { auth } from "@/auth";
import { assertCandidateAccess } from "@/lib/services/customerAccessService";
import { getCustomerContext } from "@/lib/services/customerContextService";
import { getCustomerMatrixRecords } from "@/lib/services/customerDashboardService";
import {
  getCustomerDocumentRecords,
  getCustomerNvqRecords,
} from "@/lib/services/customerPortalService";
import {
  getCustomerEusrRecords,
  getCustomerInHouseRecords,
  getCustomerNporsRecords,
  getCustomerStreetworksRecords,
} from "@/lib/services/customerTrainingRecordsService";
import { assertCompanyMatch } from "@/lib/services/securityService";
import { getWorkforceById } from "@/lib/services/workforceService";
import type { CustomerContext } from "@/types/models";

export const dynamic = "force-dynamic";

function nameKey(value: string | null | undefined): string {
  return (value ?? "").trim().toLowerCase();
}

function matchesCandidate(
  row: { candidateName?: string | null; workforceId?: string | null; candidate?: string | null },
  candidateId: string,
  candidateName: string,
): boolean {
  if (row.workforceId && row.workforceId === candidateId) return true;
  const key = nameKey(candidateName);
  if (key && nameKey(row.candidateName) === key) return true;
  if (key && nameKey(row.candidate) === key) return true;
  return false;
}

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
    const { logCandidateView } = await import("@/lib/services/auditLogService");
    await logCandidateView({
      userEmail: email,
      roleType: context.roleLabel,
      company: context.companyName,
      candidateId,
      candidateName: candidate.candidateName,
      success: false,
      errorMessage: "Candidate company mismatch",
    });
    redirect("/access-denied");
  }

  if (!assertCandidateAccess(candidate, context)) {
    const { logCandidateView } = await import("@/lib/services/auditLogService");
    await logCandidateView({
      userEmail: email,
      roleType: context.roleLabel,
      company: context.companyName,
      candidateId,
      candidateName: candidate.candidateName,
      success: false,
      errorMessage: "Candidate outside access scope",
    });
    redirect("/access-denied");
  }

  const { logCandidateView } = await import("@/lib/services/auditLogService");
  await logCandidateView({
    userEmail: email,
    roleType: context.roleLabel,
    company: context.companyName,
    candidateId,
    candidateName: candidate.candidateName,
    success: true,
  });

  const [
    matrixRows,
    npors,
    eusr,
    streetworks,
    inHouse,
    nvq,
    documents,
  ] = await Promise.all([
    getCustomerMatrixRecords(context.companyName, context),
    getCustomerNporsRecords(context.companyId, context),
    getCustomerEusrRecords(context.companyId, context),
    getCustomerStreetworksRecords(context.companyId, context),
    getCustomerInHouseRecords(context.companyId, context),
    getCustomerNvqRecords(context.companyId, context),
    getCustomerDocumentRecords(
      context.companyId,
      context.canDownload,
      context,
    ),
  ]);

  const matrixRow =
    matrixRows.find(
      (row) =>
        row.candidateId === candidateId ||
        nameKey(row.candidateName) === nameKey(candidate.candidateName),
    ) ?? null;

  const filterRows = <T extends { candidateName?: string | null; workforceId?: string | null; candidate?: string | null }>(
    rows: T[],
  ) =>
    rows.filter((row) =>
      matchesCandidate(row, candidateId, candidate.candidateName),
    );

  return (
    <CandidateProfileView
      candidate={candidate}
      matrixRow={matrixRow}
      matrixReturnHref={returnRaw}
      nporsRecords={filterRows(npors)}
      eusrRecords={filterRows(eusr)}
      streetworksRecords={filterRows(streetworks)}
      inHouseRecords={filterRows(inHouse)}
      nvqRecords={filterRows(nvq)}
      documents={filterRows(documents)}
    />
  );
}
