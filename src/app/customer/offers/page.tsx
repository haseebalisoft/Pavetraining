import { redirect } from "next/navigation";

import { OffersView } from "@/components/customer/OffersView";
import { auth } from "@/auth";
import { getCustomerContext } from "@/lib/services/customerContextService";
import { getCustomerOfferRecords } from "@/lib/services/customerPortalService";
import type { CustomerContext } from "@/types/models";

export const dynamic = "force-dynamic";

export default async function CustomerOffersPage() {
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

  const records = await getCustomerOfferRecords(context.companyId);

  return <OffersView companyName={context.companyName} records={records} />;
}
