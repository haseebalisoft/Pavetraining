import { redirect } from "next/navigation";

import { BrandLogo } from "@/components/brand/BrandLogo";
import { auth, signIn } from "@/auth";
import styles from "./login.module.css";

export const dynamic = "force-dynamic";

export default async function LoginPage() {
  const session = await auth();
  if (session?.user?.email) {
    redirect("/");
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
          Use your Microsoft work account. Access is granted from your PAVE
          permissions after sign-in.
        </p>
        <form
          action={async () => {
            "use server";
            await signIn("microsoft-entra-id", { redirectTo: "/" });
          }}
        >
          <button className={styles.button} type="submit">
            Sign in with Microsoft
          </button>
        </form>
      </section>
    </main>
  );
}
