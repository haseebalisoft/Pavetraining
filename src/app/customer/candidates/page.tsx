import { redirect } from "next/navigation";

import { CandidatesView } from "@/components/customer/CandidatesView";
import { auth } from "@/auth";
import {
  getAllowedWorkforceForCustomer,
  isCompanyWideScope,
} from "@/lib/services/customerAccessService";
import { getCustomerContext } from "@/lib/services/customerContextService";
import { getCustomerMatrixRecords } from "@/lib/services/customerDashboardService";
import { getWorkforceByCompanyName } from "@/lib/services/workforceService";
import type { CustomerContext, CustomerMatrixRecord } from "@/types/models";

export const dynamic = "force-dynamic";

function nameKey(value: string | null | undefined): string {
  return (value ?? "").trim().toLowerCase().replace(/\s+/g, " ");
}

function enrichFromMatrix(
  candidateId: string,
  candidateName: string,
  byId: Map<string, CustomerMatrixRecord>,
  byName: Map<string, CustomerMatrixRecord>,
): CustomerMatrixRecord | null {
  return (
    byId.get(candidateId) ??
    byName.get(nameKey(candidateName)) ??
    null
  );
}

function accessEmptyHintFor(context: CustomerContext): string | null {
  if (isCompanyWideScope(context.normalizedAccessScope)) return null;
  if (context.normalizedAccessScope === "Department") {
    const depts = context.departmentScopes.join(", ");
    return depts
      ? `Your Training Manager access is Department Only (${depts}). No workforce rows match those departments. Ask an admin to set Access Scope to Full Company, or align Workforce department names.`
      : "Your Training Manager access is Department Only, but no departments are assigned on your Permissions row. Ask an admin to set Access Scope to Full Company or add departments.";
  }
  if (context.normalizedAccessScope === "AssignedCandidates") {
    return "Your access is limited to assigned candidates. None are linked to you as Training Manager or Supervisor yet.";
  }
  if (context.normalizedAccessScope === "CandidateOnly") {
    return "Your access is limited to your own candidate profile.";
  }
  return null;
}

export default async function CustomerCandidatesPage() {
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

  const [candidates, matrix, companyWorkforceCount] = await Promise.all([
    getAllowedWorkforceForCustomer(context),
    getCustomerMatrixRecords(context.companyName, context),
    isCompanyWideScope(context.normalizedAccessScope)
      ? Promise.resolve(0)
      : getWorkforceByCompanyName(context.companyName).then(
          (rows) => rows.length,
        ),
  ]);

  const byId = new Map<string, CustomerMatrixRecord>();
  const byName = new Map<string, CustomerMatrixRecord>();
  for (const row of matrix) {
    if (row.candidateId?.trim()) {
      byId.set(row.candidateId.trim(), row);
    }
    const key = nameKey(row.candidateName);
    if (key && !byName.has(key)) {
      byName.set(key, row);
    }
  }

  const enriched = candidates.map((candidate) => {
    const matrixRow = enrichFromMatrix(
      candidate.id,
      candidate.candidateName,
      byId,
      byName,
    );
    return {
      ...candidate,
      cscsNumber: matrixRow?.cscsNumber ?? candidate.cscsNumber,
      cscsExpiry: matrixRow?.cscsExpiry ?? candidate.cscsExpiry,
      swqrNumber: matrixRow?.swqrNumber ?? candidate.swqrNumber,
      swqrExpiry: matrixRow?.swqrExpiry ?? candidate.swqrExpiry,
      eusrNumber: matrixRow?.eusrNumber ?? candidate.eusrNumber,
      eusrExpiry: matrixRow?.eusrExpiry ?? candidate.eusrExpiry,
      nporsNumber: matrixRow?.nporsNumber ?? candidate.nporsNumbers,
      nporsCategories: matrixRow?.nporsCategories ?? null,
      nporsExpiry: matrixRow?.nporsExpiry ?? null,
      inHouseCourse: matrixRow?.inHouseCourse ?? null,
      inHouseExpiry: matrixRow?.inHouseExpiry ?? null,
      department: matrixRow?.department ?? candidate.department,
      trainingManager:
        matrixRow?.trainingManager ?? candidate.trainingManager,
      supervisor: matrixRow?.supervisor ?? candidate.supervisor,
    };
  });

  const accessEmptyHint =
    enriched.length === 0 &&
    (companyWorkforceCount > 0 ||
      !isCompanyWideScope(context.normalizedAccessScope))
      ? accessEmptyHintFor(context)
      : null;

  return (
    <CandidatesView
      companyName={context.companyName}
      candidates={enriched}
      accessEmptyHint={accessEmptyHint}
    />
  );
}
