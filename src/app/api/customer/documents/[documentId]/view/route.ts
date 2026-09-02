import { handleApiError } from "@/lib/api/apiGuards";
import { fileResponse } from "@/lib/http/fileResponse";
import {
  logDocumentView,
  sanitizeAuditError,
} from "@/lib/services/auditLogService";
import {
  AccessDeniedError,
  NotFoundError,
} from "@/lib/services/errorHandler";
import {
  assertCustomerVisible,
  requireCustomerAccess,
} from "@/lib/services/securityService";
import {
  downloadCustomerDocumentFile,
  getCustomerDocumentForAccess,
} from "@/lib/services/customerPortalService";

export const dynamic = "force-dynamic";

/**
 * Inline preview for a customer document when:
 * - Active Customer permission
 * - document belongs to assigned company
 * - CustomerVisible = Yes
 * - item is a file (FSObjType = 0)
 *
 * Does not require CanDownload (view-only customers can still preview).
 */
export async function GET(
  request: Request,
  context: { params: Promise<{ documentId: string }> },
) {
  let email = "unknown";
  let roleType: string | null = null;
  let company: string | null = null;
  let documentId = "";
  let documentName: string | null = null;

  try {
    const params = await context.params;
    documentId = params.documentId;
    const customer = await requireCustomerAccess(request);
    email = customer.loggedInEmail;
    roleType = customer.roleLabel;
    company = customer.companyName;

    const document = await getCustomerDocumentForAccess(
      customer.companyId,
      documentId,
      customer,
    );

    if (!document) {
      throw new NotFoundError();
    }

    documentName = document.name;

    if (!document.isFile || !document.companyMatches || !document.scopeAllowed) {
      throw new AccessDeniedError();
    }

    assertCustomerVisible(document.customerVisible);

    const file = await downloadCustomerDocumentFile(documentId);
    if (!file) {
      throw new NotFoundError();
    }

    await logDocumentView({
      userEmail: email,
      roleType,
      company,
      documentId,
      documentName,
      success: true,
      request,
    });

    return fileResponse(
      file.content,
      file.fileName ?? document.name,
      file.contentType,
      "inline",
    );
  } catch (error) {
    await logDocumentView({
      userEmail: email,
      roleType,
      company,
      documentId,
      documentName,
      success: false,
      errorMessage: sanitizeAuditError(error),
      request,
    });

    return handleApiError(
      "GET /api/customer/documents/[documentId]/view",
      error,
    );
  }
}
