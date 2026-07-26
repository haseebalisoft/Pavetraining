import { NextResponse } from "next/server";

import { handleApiError } from "@/lib/api/apiGuards";
import { writeAuditLog } from "@/lib/services/auditLogService";
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
 * Downloads a customer document only when:
 * - Active Customer permission
 * - CanDownload = true
 * - document belongs to assigned company
 * - CustomerVisible = Yes
 * - item is a file (FSObjType = 0)
 */
export async function GET(
  request: Request,
  context: { params: Promise<{ documentId: string }> },
) {
  let email = "unknown";
  let documentId = "";

  try {
    const params = await context.params;
    documentId = params.documentId;
    const customer = await requireCustomerAccess(request);
    email = customer.loggedInEmail;

    if (!customer.canDownload) {
      throw new AccessDeniedError();
    }

    const document = await getCustomerDocumentForAccess(
      customer.companyId,
      documentId,
      customer,
    );

    if (!document) {
      throw new NotFoundError();
    }

    if (!document.isFile || !document.companyMatches || !document.scopeAllowed) {
      throw new AccessDeniedError();
    }

    assertCustomerVisible(document.customerVisible);

    const file = await downloadCustomerDocumentFile(documentId);
    if (!file) {
      throw new NotFoundError();
    }

    await writeAuditLog({
      userEmail: email,
      action: "DOWNLOAD",
      entityName: "Customer Documents",
      itemId: documentId,
      success: true,
    });

    const fileName = file.fileName ?? document.name;
    return new NextResponse(file.content, {
      status: 200,
      headers: {
        "Content-Type": file.contentType,
        "Content-Disposition": `attachment; filename="${fileName.replace(/"/g, "")}"`,
        "Cache-Control": "no-store",
      },
    });
  } catch (error) {
    await writeAuditLog({
      userEmail: email,
      action: "DOWNLOAD",
      entityName: "Customer Documents",
      itemId: documentId || null,
      success: false,
      errorMessage: error instanceof Error ? error.message : "Unknown error",
    });

    return handleApiError(
      "GET /api/customer/documents/[documentId]/download",
      error,
    );
  }
}
