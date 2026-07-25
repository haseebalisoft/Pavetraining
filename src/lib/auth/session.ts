import "server-only";

import { auth } from "@/auth";
import { UnauthorizedError } from "@/lib/services/customerContextService";

export async function requireAuthenticatedEmail(): Promise<string> {
  const session = await auth();
  const email = session?.user?.email?.trim().toLowerCase();

  if (!email) {
    throw new UnauthorizedError();
  }

  return email;
}
