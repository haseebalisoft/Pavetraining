import { withCustomerApi } from "@/lib/api/customerApi";
import { getCustomerOfferRecords } from "@/lib/services/customerPortalService";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  return withCustomerApi(
    "GET /api/customer/offers",
    async (context) => {
      const records = await getCustomerOfferRecords(context.companyId);
      return { records };
    },
    { entityName: "Offers" },
    request,
  );
}
