import { redirect } from "next/navigation";
import type { ReactNode } from "react";

import { CustomerTopNav } from "@/components/customer/CustomerTopNav";
import { auth, signOut } from "@/auth";
import { accessScopeBadgeLabel } from "@/lib/services/customerAccessService";
import { getCustomerContext } from "@/lib/services/customerContextService";
import { getAssignedManagersForCustomer } from "@/lib/services/workforceService";
import type { CustomerContext } from "@/types/models";

import styles from "../../components/customer/customer.module.css";

export const dynamic = "force-dynamic";

async function signOutAction() {
  "use server";
  await signOut({ redirectTo: "/login" });
}

export default async function CustomerLayout({
  children,
}: Readonly<{
  children: ReactNode;
}>) {
  const session = await auth();
  const email = session?.user?.email?.trim().toLowerCase();

  if (!email) {
    redirect("/login");
  }

  let context: CustomerContext;
  try {
    context = await getCustomerContext(email);
  } catch {
    try {
      const { logAccessDenied } = await import("@/lib/services/auditLogService");
      await logAccessDenied({
        userEmail: email,
        entityType: "Customer Portal",
        entityName: "/customer",
        errorMessage: "Customer access denied",
      });
    } catch {
      // ignore audit failures
    }
    redirect("/access-denied");
  }

  const assigned = await getAssignedManagersForCustomer({
    companyName: context.companyName,
    email: context.loggedInEmail,
    displayName: context.candidateScopeName,
    customerRole: context.customerRole,
  });

  return (
    <div className={styles.shell}>
      <CustomerTopNav
        email={context.loggedInEmail}
        companyName={context.companyName}
        roleLabel={context.roleLabel}
        accessLabel={accessScopeBadgeLabel(context)}
        canDownload={context.canDownload}
        trainingManager={assigned.trainingManager}
        supervisor={assigned.supervisor}
        hideCandidatesNav={context.normalizedAccessScope === "CandidateOnly"}
        signOutAction={signOutAction}
      />
      <main className={styles.main}>{children}</main>
    </div>
  );
}
