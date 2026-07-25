import { redirect } from "next/navigation";

import { CandidatesView } from "@/components/customer/CandidatesView";
import { auth } from "@/auth";
import { getCustomerContext } from "@/lib/services/customerContextService";
import { getWorkforceByCompanyName } from "@/lib/services/workforceService";
import type { CustomerContext } from "@/types/models";

export const dynamic = "force-dynamic";

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

  const candidates = await getWorkforceByCompanyName(context.companyName);

  return (
    <CandidatesView
      companyName={context.companyName}
      candidates={candidates}
    />
  );
}
