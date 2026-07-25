import { handleApiError } from "@/lib/api/apiGuards";
import { requireAuthenticatedEmail } from "@/lib/auth/session";
import { AccessDeniedError } from "@/lib/services/errorHandler";
import { getMeContext } from "@/lib/services/customerContextService";

export const dynamic = "force-dynamic";

export async function GET() {
  try {
    const email = await requireAuthenticatedEmail();
    const me = await getMeContext(email);

    if (!me) {
      throw new AccessDeniedError();
    }

    return Response.json(me);
  } catch (error) {
    return handleApiError("GET /api/me", error);
  }
}
