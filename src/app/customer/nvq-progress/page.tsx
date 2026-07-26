import { redirect } from "next/navigation";

import { NvqProgressView } from "@/components/customer/NvqProgressView";
import { auth } from "@/auth";
import { getCustomerContext } from "@/lib/services/customerContextService";
import { getCustomerNvqRecords } from "@/lib/services/customerPortalService";
import type { CustomerContext } from "@/types/models";

export const dynamic = "force-dynamic";

export default async function CustomerNvqProgressPage() {
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

  const records = await getCustomerNvqRecords(context.companyId, context);

  return (
    <NvqProgressView companyName={context.companyName} records={records} />
  );
}
