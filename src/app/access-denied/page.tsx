import Link from "next/link";

import { BrandLogo } from "@/components/brand/BrandLogo";
import { auth, signOut } from "@/auth";
import styles from "./access-denied.module.css";

export const dynamic = "force-dynamic";

function titleAndCopyFor(
  reason: string | undefined,
  email: string | null,
): { title: string; copy: string } {
  if (reason === "not_found") {
    return {
      title: "Permission not configured",
      copy: "Your account is signed in, but no Permissions record has been set up for this email yet. Ask your administrator to add one.",
    };
  }
  if (reason === "inactive") {
    return {
      title: "Access pending",
      copy: "Your account is signed in, and a Permissions record exists, but it is not Active yet. Ask your administrator to activate it.",
    };
  }
  return {
    title: "Access denied",
    copy: email
      ? "Your account is signed in, but it does not have permission to open this area of the portal. Contact your training manager if you believe this is a mistake."
      : "You do not have access to this portal. Sign in with an authorised Microsoft account linked in Permissions.",
  };
}

export default async function AccessDeniedPage({
  searchParams,
}: {
  searchParams: Promise<{ reason?: string }>;
}) {
  const session = await auth();
  const email = session?.user?.email ?? null;
  const { reason } = await searchParams;
  const { title, copy } = titleAndCopyFor(reason, email);

  return (
    <main className={styles.page}>
      <section className={styles.panel}>
        <div className={styles.logoWrap}>
          <BrandLogo variant="compact" />
        </div>
        <h1 className={styles.title}>{title}</h1>
        <p className={styles.copy}>{copy}</p>
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
