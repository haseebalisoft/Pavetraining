import { redirect } from "next/navigation";

import { InHouseRecordsView } from "@/components/customer/InHouseRecordsView";
import { auth } from "@/auth";
import { getCustomerContext } from "@/lib/services/customerContextService";
import { getCustomerInHouseRecords } from "@/lib/services/customerTrainingRecordsService";
import type { CustomerContext } from "@/types/models";

export const dynamic = "force-dynamic";

export default async function CustomerInHouseRecordsPage() {
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

  const records = await getCustomerInHouseRecords(context.companyId);

  return (
    <InHouseRecordsView companyName={context.companyName} records={records} />
  );
}
