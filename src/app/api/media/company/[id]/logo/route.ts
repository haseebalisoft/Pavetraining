import { NextResponse } from "next/server";

import { handleApiError } from "@/lib/api/apiGuards";
import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import { NotFoundError, AccessDeniedError } from "@/lib/services/errorHandler";
import { fetchThumbnailContent } from "@/lib/services/listThumbnailService";
import { getListItemByKey } from "@/lib/services/sharePointListService";
import {
  requireAdminAccess,
  requireCustomerAccess,
} from "@/lib/services/securityService";

export const dynamic = "force-dynamic";

export async function GET(
  request: Request,
  context: { params: Promise<{ id: string }> },
) {
  try {
    const { id } = await context.params;
    let allowed = false;
    try {
      await requireAdminAccess();
      allowed = true;
    } catch {
      const customer = await requireCustomerAccess(request);
      allowed = customer.companyId === id;
    }
    if (!allowed) throw new AccessDeniedError();

    const item = await getListItemByKey("company", id);
    if (!item) throw new NotFoundError("Company not found.");
    const fields = getSharePointFields("company");
    const content = await fetchThumbnailContent(item.fields[fields.companyLogo]);
    if (!content) throw new NotFoundError("Logo not found.");

    return new NextResponse(content.bytes, {
      headers: {
        "Content-Type": content.contentType,
        "Cache-Control": "private, max-age=300",
      },
    });
  } catch (error) {
    return handleApiError("GET /api/media/company/[id]/logo", error);
  }
}
