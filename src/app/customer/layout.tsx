import { redirect } from "next/navigation";
import type { ReactNode } from "react";

import { CustomerSidebar } from "@/components/customer/CustomerSidebar";
import { auth, signOut } from "@/auth";
import { accessScopeBadgeLabel } from "@/lib/services/customerAccessService";
import { getCustomerContext } from "@/lib/services/customerContextService";
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
    redirect("/access-denied");
  }

  return (
    <div className={styles.shell}>
      <CustomerSidebar
        email={context.loggedInEmail}
        companyName={context.companyName}
        roleLabel={context.roleLabel}
        accessLabel={accessScopeBadgeLabel(context)}
        canDownload={context.canDownload}
        signOutAction={signOutAction}
      />
      <main className={styles.main}>{children}</main>
    </div>
  );
}
