import { Suspense } from "react";
import { redirect } from "next/navigation";

import { BrandLogo } from "@/components/brand/BrandLogo";
import { auth, signIn } from "@/auth";
import { LoginClient } from "./LoginClient";
import styles from "./login.module.css";

export const dynamic = "force-dynamic";

export default async function LoginPage() {
  const session = await auth();
  if (session?.user?.email) {
    redirect("/");
  }

  async function microsoftSignIn() {
    "use server";
    await signIn("microsoft-entra-id", { redirectTo: "/" });
  }

  return (
    <main className={styles.page}>
      <section className={styles.panel}>
        <div className={styles.logoWrap}>
          <BrandLogo variant="full" priority />
        </div>
        <p className={styles.tagline}>Paving the way in industry</p>
        <h1 className={styles.title}>Sign in to your portal</h1>
        <p className={styles.copy}>
          Prefer Microsoft when you can. Customers without Microsoft can use an
          email one-time code. Access still comes from Permissions.
        </p>
        <Suspense fallback={<p className={styles.copy}>Loading sign-in…</p>}>
          <LoginClient
            microsoftButton={
              <form action={microsoftSignIn}>
                <button className={styles.button} type="submit">
                  Sign in with Microsoft
                </button>
              </form>
            }
          />
        </Suspense>
      </section>
    </main>
  );
}
