import { redirect } from "next/navigation";

import { StreetworksRecordsView } from "@/components/customer/StreetworksRecordsView";
import { auth } from "@/auth";
import { getCustomerContext } from "@/lib/services/customerContextService";
import { getCustomerStreetworksRecords } from "@/lib/services/customerTrainingRecordsService";
import type { CustomerContext } from "@/types/models";

export const dynamic = "force-dynamic";

export default async function CustomerStreetworksRecordsPage() {
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

  const records = await getCustomerStreetworksRecords(context.companyId);

  return (
    <StreetworksRecordsView
      companyName={context.companyName}
      records={records}
    />
  );
}
