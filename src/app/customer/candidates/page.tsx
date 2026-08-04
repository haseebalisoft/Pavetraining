import { redirect } from "next/navigation";

import { CandidatesView } from "@/components/customer/CandidatesView";
import { auth } from "@/auth";
import { getAllowedWorkforceForCustomer } from "@/lib/services/customerAccessService";
import { getCustomerContext } from "@/lib/services/customerContextService";
import { getCustomerMatrixRecords } from "@/lib/services/customerDashboardService";
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

  const [candidates, matrix] = await Promise.all([
    getAllowedWorkforceForCustomer(context),
    getCustomerMatrixRecords(context.companyName, context),
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

  return (
    <CandidatesView
      companyName={context.companyName}
      candidates={enriched}
    />
  );
}
