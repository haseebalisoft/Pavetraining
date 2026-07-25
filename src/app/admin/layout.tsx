import { redirect } from "next/navigation";
import { Suspense, type ReactNode } from "react";

import { AdminProviders } from "@/components/admin/AdminProviders";
import { AdminSidebar } from "@/components/admin/AdminSidebar";
import { LoadingState } from "@/components/ui/States";
import { auth, signOut } from "@/auth";
import { getAdminContext } from "@/lib/services/customerContextService";
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
    redirect("/access-denied");
  }

  return (
    <div className={styles.shell}>
      <AdminSidebar email={context.loggedInEmail} signOutAction={signOutAction} />
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
