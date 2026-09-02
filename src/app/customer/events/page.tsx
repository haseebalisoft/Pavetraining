import { redirect } from "next/navigation";

import { EventsView } from "@/components/customer/EventsView";
import { auth } from "@/auth";
import { getCustomerContext } from "@/lib/services/customerContextService";
import {
  getCustomerDocumentRecords,
  getCustomerEventRecords,
} from "@/lib/services/customerPortalService";
import type { CustomerContext } from "@/types/models";

export const dynamic = "force-dynamic";

export default async function CustomerEventsPage({
  searchParams,
}: {
  searchParams: Promise<{ event?: string | string[] }>;
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

  const query = await searchParams;
  const eventRaw = Array.isArray(query.event) ? query.event[0] : query.event;
  const initialEventId = eventRaw?.trim() || null;

  const [records, documents] = await Promise.all([
    getCustomerEventRecords(context.companyId),
    getCustomerDocumentRecords(
      context.companyId,
      context.canDownload,
      context,
    ).catch(() => []),
  ]);

  return (
    <EventsView
      companyName={context.companyName}
      records={records}
      documents={documents}
      initialEventId={initialEventId}
    />
  );
}
