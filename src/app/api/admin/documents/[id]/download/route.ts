import { NextResponse } from "next/server";

import { handleApiError } from "@/lib/api/apiGuards";
import { writeAuditLog } from "@/lib/services/auditLogService";
import { NotFoundError } from "@/lib/services/errorHandler";
import { requireAdminAccess } from "@/lib/services/securityService";
import {
  getListItemByKey,
  getListItemFileContent,
} from "@/lib/services/sharePointListService";

export const dynamic = "force-dynamic";

export async function GET(
  request: Request,
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
      action: "DOWNLOAD",
      entityName: "Customer Documents",
      itemId: id,
      success: true,
    });

    const dispositionParam = new URL(request.url).searchParams.get(
      "disposition",
    );
    const disposition =
      dispositionParam === "inline" ? "inline" : "attachment";
    const fileName = (file.fileName ?? "document").replace(/"/g, "");

    return new NextResponse(file.content, {
      status: 200,
      headers: {
        "Content-Type": file.contentType,
        "Content-Disposition": `${disposition}; filename="${fileName}"`,
        "Cache-Control": "no-store",
      },
    });
  } catch (error) {
    await writeAuditLog({
      userEmail: email,
      action: "DOWNLOAD",
      entityName: "Customer Documents",
      itemId: id || null,
      success: false,
      errorMessage: error instanceof Error ? error.message : "Unknown error",
    });

    return handleApiError(
      "GET /api/admin/documents/[id]/download",
      error,
    );
  }
}
