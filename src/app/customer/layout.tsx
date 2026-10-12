import Link from "next/link";
import { redirect } from "next/navigation";
import type { ReactNode } from "react";

import { BrandLogo } from "@/components/brand/BrandLogo";
import { CustomerTopNav } from "@/components/customer/CustomerTopNav";
import { auth, signOut } from "@/auth";
import { getCompanyById } from "@/lib/services/companyService";
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

  const [assigned, company] = await Promise.all([
    getAssignedManagersForCustomer({
      companyName: context.companyName,
      email: context.loggedInEmail,
      displayName: context.candidateScopeName,
      customerRole: context.customerRole,
    }),
    // Same cached read the dashboard uses — only for the nav logo.
    getCompanyById(context.companyId).catch(() => null),
  ]);

  return (
    <div className={styles.shell}>
      <CustomerTopNav
        email={context.loggedInEmail}
        companyName={context.companyName}
        companyLogo={company?.companyLogo ?? null}
        roleLabel={context.roleLabel}
        accessLabel={accessScopeBadgeLabel(context)}
        canDownload={context.canDownload}
        trainingManager={assigned.trainingManager}
        supervisor={assigned.supervisor}
        hideCandidatesNav={context.normalizedAccessScope === "CandidateOnly"}
        signOutAction={signOutAction}
      />
      <main className={styles.main}>{children}</main>
      <footer className={styles.portalFooter}>
        <div className={styles.portalFooterInner}>
          <span className={styles.portalFooterBrand}>
            <span className={styles.topNavBrandMark}>
              <BrandLogo variant="mark" />
            </span>
            © {new Date().getFullYear()} PAVE Training. All rights reserved.
          </span>
          <nav className={styles.portalFooterLinks} aria-label="Footer">
            <Link href="/customer/support">Support</Link>
          </nav>
        </div>
      </footer>
    </div>
  );
}
