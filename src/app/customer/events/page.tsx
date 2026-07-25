import { redirect } from "next/navigation";

import { EventsView } from "@/components/customer/EventsView";
import { auth } from "@/auth";
import { getCustomerContext } from "@/lib/services/customerContextService";
import { getCustomerEventRecords } from "@/lib/services/customerPortalService";
import type { CustomerContext } from "@/types/models";

export const dynamic = "force-dynamic";

export default async function CustomerEventsPage() {
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

  const records = await getCustomerEventRecords(context.companyId);

  return <EventsView companyName={context.companyName} records={records} />;
}
