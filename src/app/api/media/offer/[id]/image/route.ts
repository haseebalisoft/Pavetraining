import { NextResponse } from "next/server";

import { handleApiError } from "@/lib/api/apiGuards";
import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import {
  AccessDeniedError,
  NotFoundError,
} from "@/lib/services/errorHandler";
import { fetchThumbnailContent } from "@/lib/services/listThumbnailService";
import { getCustomerOfferRecords } from "@/lib/services/customerPortalService";
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
      const offers = await getCustomerOfferRecords(customer.companyId);
      allowed = offers.some((offer) => offer.id === id);
    }

    if (!allowed) throw new AccessDeniedError();

    const item = await getListItemByKey("offersPromotions", id);
    if (!item) throw new NotFoundError("Offer not found.");

    const fields = getSharePointFields("offersPromotions");
    const content = await fetchThumbnailContent(item.fields[fields.image]);
    if (!content) throw new NotFoundError("Offer image not found.");

    return new NextResponse(content.bytes, {
      headers: {
        "Content-Type": content.contentType,
        "Cache-Control": "private, max-age=300",
      },
    });
  } catch (error) {
    return handleApiError("GET /api/media/offer/[id]/image", error);
  }
}
