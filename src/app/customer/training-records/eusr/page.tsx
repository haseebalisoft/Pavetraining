import { redirect } from "next/navigation";

import { EusrRecordsView } from "@/components/customer/EusrRecordsView";
import { auth } from "@/auth";
import { getCustomerContext } from "@/lib/services/customerContextService";
import { getCustomerEusrRecords } from "@/lib/services/customerTrainingRecordsService";
import type { CustomerContext } from "@/types/models";

export const dynamic = "force-dynamic";

export default async function CustomerEusrRecordsPage() {
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

  const records = await getCustomerEusrRecords(context.companyId, context);

  return (
    <EusrRecordsView companyName={context.companyName} records={records} />
  );
}
