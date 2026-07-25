import { redirect } from "next/navigation";

import { NporsRecordsView } from "@/components/customer/NporsRecordsView";
import { auth } from "@/auth";
import { getCustomerContext } from "@/lib/services/customerContextService";
import { getCustomerNporsRecords } from "@/lib/services/customerTrainingRecordsService";
import type { CustomerContext } from "@/types/models";

export const dynamic = "force-dynamic";

export default async function CustomerNporsRecordsPage() {
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

  const records = await getCustomerNporsRecords(context.companyId);

  return (
    <NporsRecordsView companyName={context.companyName} records={records} />
  );
}
