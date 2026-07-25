import { redirect } from "next/navigation";

import { CandidateProfileView } from "@/components/customer/CandidateProfileView";
import { auth } from "@/auth";
import { getCustomerContext } from "@/lib/services/customerContextService";
import { assertCompanyMatch } from "@/lib/services/securityService";
import { getWorkforceById } from "@/lib/services/workforceService";
import type { CustomerContext } from "@/types/models";

export const dynamic = "force-dynamic";

export default async function CustomerCandidateProfilePage({
  params,
}: {
  params: Promise<{ candidateId: string }>;
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
  const candidate = await getWorkforceById(candidateId);

  if (!candidate) {
    redirect("/customer/candidates");
  }

  try {
    assertCompanyMatch(candidate.companyName, context.companyName);
  } catch {
    redirect("/access-denied");
  }

  return <CandidateProfileView candidate={candidate} />;
}
