import { redirect } from "next/navigation";

import { TrainingMatrixView } from "@/components/customer/TrainingMatrixView";
import { auth } from "@/auth";
import { getCustomerContext } from "@/lib/services/customerContextService";
import { getCustomerMatrixRecords } from "@/lib/services/customerDashboardService";
import type { CustomerContext } from "@/types/models";

export const dynamic = "force-dynamic";

type MatrixFilter = "all" | "review" | "expired" | "expiring";

function parseFilter(value: string | string[] | undefined): MatrixFilter {
  const raw = Array.isArray(value) ? value[0] : value;
  if (
    raw === "review" ||
    raw === "expired" ||
    raw === "expiring" ||
    raw === "all"
  ) {
    return raw;
  }
  return "all";
}

export default async function CustomerTrainingMatrixPage({
  searchParams,
}: {
  searchParams: Promise<{ filter?: string | string[] }>;
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

  const params = await searchParams;
  const records = await getCustomerMatrixRecords(context.companyName);

  return (
    <TrainingMatrixView
      companyName={context.companyName}
      records={records}
      initialFilter={parseFilter(params.filter)}
    />
  );
}
