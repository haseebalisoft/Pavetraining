import Link from "next/link";

import { BrandLogo } from "@/components/brand/BrandLogo";
import { auth, signOut } from "@/auth";
import styles from "./access-denied.module.css";

export const dynamic = "force-dynamic";

export default async function AccessDeniedPage() {
  const session = await auth();
  const email = session?.user?.email ?? null;

  return (
    <main className={styles.page}>
      <section className={styles.panel}>
        <div className={styles.logoWrap}>
          <BrandLogo variant="compact" />
        </div>
        <h1 className={styles.title}>Access denied</h1>
        <p className={styles.copy}>
          {email
            ? "Your account is signed in, but it does not have permission to open this area of the portal. Contact your training manager if you believe this is a mistake."
            : "You do not have access to this portal. Sign in with an authorised Microsoft account linked in Permissions."}
        </p>
        <div className={styles.actions}>
          <Link className={styles.secondary} href="/login">
            Back to login
          </Link>
          {email ? (
            <form
              action={async () => {
                "use server";
                await signOut({ redirectTo: "/login" });
              }}
            >
              <button className={styles.primary} type="submit">
                Sign out
              </button>
            </form>
          ) : null}
        </div>
      </section>
    </main>
  );
}
