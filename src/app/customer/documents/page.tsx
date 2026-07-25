import { redirect } from "next/navigation";

import { DocumentsView } from "@/components/customer/DocumentsView";
import { auth } from "@/auth";
import { getCustomerContext } from "@/lib/services/customerContextService";
import { getCustomerDocumentRecords } from "@/lib/services/customerPortalService";
import type { CustomerContext } from "@/types/models";

export const dynamic = "force-dynamic";

export default async function CustomerDocumentsPage() {
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

  const records = await getCustomerDocumentRecords(
    context.companyId,
    context.canDownload,
  );

  return (
    <DocumentsView
      companyName={context.companyName}
      canDownload={context.canDownload}
      records={records}
    />
  );
}
