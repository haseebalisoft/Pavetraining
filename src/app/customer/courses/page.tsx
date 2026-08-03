import { redirect } from "next/navigation";

import { CustomerCoursesView } from "@/components/customer/CustomerCoursesView";
import { auth } from "@/auth";
import { getCustomerCourseCatalogue } from "@/lib/services/customerCoursesService";
import { getCustomerContext } from "@/lib/services/customerContextService";
import { getCustomerDocumentRecords } from "@/lib/services/customerPortalService";
import type { CustomerContext } from "@/types/models";

export const dynamic = "force-dynamic";

export default async function CustomerCoursesPage() {
  const session = await auth();
  const email = session?.user?.email?.trim().toLowerCase();
  if (!email) redirect("/login");

  let context: CustomerContext;
  try {
    context = await getCustomerContext(email);
  } catch {
    redirect("/access-denied");
  }

  const [courses, documents] = await Promise.all([
    getCustomerCourseCatalogue(),
    getCustomerDocumentRecords(
      context.companyId,
      context.canDownload,
      context,
    ),
  ]);

  return (
    <CustomerCoursesView
      companyName={context.companyName}
      courses={courses}
      resourceDocs={documents}
    />
  );
}
