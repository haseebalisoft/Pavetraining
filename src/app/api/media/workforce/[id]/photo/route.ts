import { NextResponse } from "next/server";

import { handleApiError } from "@/lib/api/apiGuards";
import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import { assertCandidateAccess } from "@/lib/services/customerAccessService";
import { NotFoundError, AccessDeniedError } from "@/lib/services/errorHandler";
import { fetchThumbnailContent } from "@/lib/services/listThumbnailService";
import { getListItemByKey } from "@/lib/services/sharePointListService";
import {
  assertCompanyMatch,
  requireAdminAccess,
  requireCustomerAccess,
} from "@/lib/services/securityService";
import { getWorkforceById } from "@/lib/services/workforceService";

export const dynamic = "force-dynamic";

export async function GET(
  request: Request,
  context: { params: Promise<{ id: string }> },
) {
  try {
    const { id } = await context.params;
    try {
      await requireAdminAccess();
    } catch {
      const customer = await requireCustomerAccess(request);
      const candidate = await getWorkforceById(id);
      if (!candidate) throw new NotFoundError("Candidate not found.");
      try {
        assertCompanyMatch(candidate.companyName, customer.companyName);
      } catch {
        throw new AccessDeniedError();
      }
      if (!assertCandidateAccess(candidate, customer)) {
        throw new AccessDeniedError();
      }
    }

    const item = await getListItemByKey("workforce", id);
    if (!item) throw new NotFoundError("Candidate not found.");
    const fields = getSharePointFields("workforce");
    const content = await fetchThumbnailContent(item.fields[fields.photo]);
    if (!content) throw new NotFoundError("Photo not found.");

    return new NextResponse(content.bytes, {
      headers: {
        "Content-Type": content.contentType,
        "Cache-Control": "private, max-age=300",
      },
    });
  } catch (error) {
    return handleApiError("GET /api/media/workforce/[id]/photo", error);
  }
}
