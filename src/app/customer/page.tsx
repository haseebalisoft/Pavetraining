import { Suspense } from "react";
import { redirect } from "next/navigation";

import { TrainingMatrixView } from "@/components/customer/TrainingMatrixView";
import { LoadingState } from "@/components/ui/States";
import { auth } from "@/auth";
import { getCustomerContext } from "@/lib/services/customerContextService";
import { getCustomerMatrixRecords } from "@/lib/services/customerDashboardService";
import type { CustomerContext } from "@/types/models";

export const dynamic = "force-dynamic";

export default async function CustomerPage() {
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

  const records = await getCustomerMatrixRecords(
    context.companyName,
    context,
  );

  return (
    <Suspense fallback={<LoadingState label="Loading training matrix…" />}>
      <TrainingMatrixView
        companyName={context.companyName}
        records={records}
        isLanding
      />
    </Suspense>
  );
}
