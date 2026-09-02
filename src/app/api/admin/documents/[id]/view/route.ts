import { handleApiError } from "@/lib/api/apiGuards";
import { fileResponse } from "@/lib/http/fileResponse";
import { writeAuditLog } from "@/lib/services/auditLogService";
import { NotFoundError } from "@/lib/services/errorHandler";
import { requireAdminAccess } from "@/lib/services/securityService";
import {
  getListItemByKey,
  getListItemFileContent,
} from "@/lib/services/sharePointListService";

export const dynamic = "force-dynamic";

/**
 * Inline preview. Kept on a /view URL (not /download) so browsers open PDFs
 * and images instead of saving them.
 */
export async function GET(
  _request: Request,
  context: { params: Promise<{ id: string }> },
) {
  let email = "unknown";
  let id = "";

  try {
    const params = await context.params;
    id = params.id;
    const admin = await requireAdminAccess();
    email = admin.loggedInEmail;

    const item = await getListItemByKey("customerDocuments", id);
    if (!item) {
      throw new NotFoundError();
    }

    const file = await getListItemFileContent("customerDocuments", id);
    if (!file) {
      throw new NotFoundError();
    }

    await writeAuditLog({
      userEmail: email,
      action: "VIEW",
      entityName: "Customer Documents",
      itemId: id,
      success: true,
    });

    return fileResponse(
      file.content,
      file.fileName ?? "document",
      file.contentType,
      "inline",
    );
  } catch (error) {
    await writeAuditLog({
      userEmail: email,
      action: "VIEW",
      entityName: "Customer Documents",
      itemId: id || null,
      success: false,
      errorMessage: error instanceof Error ? error.message : "Unknown error",
    });

    return handleApiError("GET /api/admin/documents/[id]/view", error);
  }
}
