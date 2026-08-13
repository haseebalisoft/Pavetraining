import { redirect } from "next/navigation";
import { Suspense, type ReactNode } from "react";

import { AdminProviders } from "@/components/admin/AdminProviders";
import { AdminTopNav } from "@/components/admin/AdminTopNav";
import { allowedAdminNavHrefs } from "@/components/admin/adminNavItems";
import { LoadingState } from "@/components/ui/States";
import { auth, signOut } from "@/auth";
import { getAdminContext } from "@/lib/services/customerContextService";
import { getActivePermissionByEmail } from "@/lib/services/permissionService";
import type { AdminContext } from "@/types/models";

import styles from "../../components/admin/admin.module.css";

export const dynamic = "force-dynamic";

async function signOutAction() {
  "use server";
  await signOut({ redirectTo: "/login" });
}

export default async function AdminLayout({
  children,
}: Readonly<{
  children: ReactNode;
}>) {
  const session = await auth();
  const email = session?.user?.email?.trim().toLowerCase();

  if (!email) {
    redirect("/login");
  }

  let context: AdminContext;
  try {
    context = await getAdminContext(email);
  } catch {
    // Non-admin who typed /admin in the URL. If they have a customer-side
    // role (Training Manager / Supervisor / Candidate), send them home to
    // /customer/dashboard rather than the generic /access-denied — matches
    // the client's strict role rules. Only truly unpermissioned accounts
    // fall through to /access-denied.
    let customerFallback: string | null = null;
    try {
      const permission = await getActivePermissionByEmail(email);
      if (permission?.canAccessCustomer) {
        customerFallback = "/customer/dashboard";
      }
    } catch {
      // ignore lookup failure — treat as access denied.
    }
    try {
      const { logAccessDenied } = await import("@/lib/services/auditLogService");
      await logAccessDenied({
        userEmail: email,
        entityType: "Admin Portal",
        entityName: "/admin",
        errorMessage: customerFallback
          ? "Non-admin role redirected to customer portal"
          : "Admin access denied",
      });
    } catch {
      // ignore audit failures
    }
    redirect(customerFallback ?? "/access-denied");
  }

  // Server-side diagnostic: never leaves the Node process, but makes it easy
  // to grep the deployment logs when a user reports "Bulk Upload is missing".
  const allowedNavHrefs = allowedAdminNavHrefs({
    isSharePointAdmin: context.isSharePointAdmin,
  });
  console.debug(
    "[admin/layout] resolved role",
    JSON.stringify({
      email: context.loggedInEmail,
      roleLabel: context.roleLabel,
      sharePointRoleType: context.sharePointRoleType,
      customerRole: context.customerRole,
      isSharePointAdmin: context.isSharePointAdmin,
      isAlwaysAdminEmail: context.isAlwaysAdminEmail,
      allowedNavHrefs,
    }),
  );

  return (
    <div className={styles.shell}>
      <AdminTopNav
        email={context.loggedInEmail}
        roleLabel={context.roleLabel}
        isSharePointAdmin={context.isSharePointAdmin}
        isAlwaysAdminEmail={context.isAlwaysAdminEmail}
        sharePointRoleType={context.sharePointRoleType}
        customerRole={context.customerRole}
        signOutAction={signOutAction}
      />
      <main className={styles.main}>
        <AdminProviders>
          <Suspense fallback={<LoadingState label="Loading admin…" />}>
            {children}
          </Suspense>
        </AdminProviders>
      </main>
    </div>
  );
}
