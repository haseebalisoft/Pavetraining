import { handleApiError } from "@/lib/api/apiGuards";
import { requireAuthenticatedEmail } from "@/lib/auth/session";
import { logLogin, sanitizeAuditError } from "@/lib/services/auditLogService";
import { AccessDeniedError } from "@/lib/services/errorHandler";
import { getMeContext } from "@/lib/services/customerContextService";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  let email = "unknown";
  try {
    email = await requireAuthenticatedEmail();
    const me = await getMeContext(email);

    if (!me) {
      await logLogin({
        userEmail: email,
        success: false,
        errorMessage: "No active permission for this account.",
        request,
      });
      throw new AccessDeniedError();
    }

    return Response.json(me);
  } catch (error) {
    if (!(error instanceof AccessDeniedError)) {
      await logLogin({
        userEmail: email,
        success: false,
        errorMessage: sanitizeAuditError(error),
        request,
      });
    }
    return handleApiError("GET /api/me", error);
  }
}
