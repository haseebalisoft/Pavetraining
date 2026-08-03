import { redirect } from "next/navigation";

import { CustomerSupportView } from "@/components/customer/CustomerSupportView";
import { auth } from "@/auth";
import { getCustomerContext } from "@/lib/services/customerContextService";
import type { CustomerContext } from "@/types/models";

export const dynamic = "force-dynamic";

export default async function CustomerSupportPage() {
  const session = await auth();
  const email = session?.user?.email?.trim().toLowerCase();
  if (!email) redirect("/login");

  let context: CustomerContext;
  try {
    context = await getCustomerContext(email);
  } catch {
    redirect("/access-denied");
  }

  return (
    <CustomerSupportView
      companyName={context.companyName}
      defaultEmail={context.loggedInEmail}
    />
  );
}
