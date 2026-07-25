import { redirect } from "next/navigation";
import { isRedirectError } from "next/dist/client/components/redirect-error";

import { auth } from "@/auth";
import { getMeContext } from "@/lib/services/customerContextService";

export const dynamic = "force-dynamic";

export default async function HomePage() {
  const session = await auth();
  const email = session?.user?.email?.trim().toLowerCase();

  if (!email) {
    redirect("/login");
  }

  try {
    const me = await getMeContext(email);

    if (!me) {
      redirect("/access-denied");
    }

    redirect(me.redirectTo);
  } catch (error) {
    if (isRedirectError(error)) {
      throw error;
    }

    console.error("[HomePage] Failed to resolve permission route", error);
    redirect("/access-denied");
  }
}
