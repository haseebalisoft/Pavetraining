import { Suspense } from "react";
import { redirect } from "next/navigation";

import { DocumentsView } from "@/components/customer/DocumentsView";
import { LoadingState } from "@/components/ui/States";
import { auth } from "@/auth";
import { getCustomerContext } from "@/lib/services/customerContextService";
import { getCustomerDocumentRecords } from "@/lib/services/customerPortalService";
import type { CustomerContext } from "@/types/models";

export const dynamic = "force-dynamic";

async function DocumentsContent({ context }: { context: CustomerContext }) {
  const records = await getCustomerDocumentRecords(
    context.companyId,
    context.canDownload,
    context,
  );

  return (
    <DocumentsView
      companyName={context.companyName}
      email={context.loggedInEmail}
      canDownload={context.canDownload}
      records={records}
    />
  );
}

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

  return (
    <Suspense fallback={<LoadingState label="Loading documents…" />}>
      <DocumentsContent context={context} />
    </Suspense>
  );
}
