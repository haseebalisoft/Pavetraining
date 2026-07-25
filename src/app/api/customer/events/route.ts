import { withCustomerApi } from "@/lib/api/customerApi";
import { getCustomerEventRecords } from "@/lib/services/customerPortalService";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  return withCustomerApi(
    "GET /api/customer/events",
    async (context) => {
      const records = await getCustomerEventRecords(context.companyId);
      return { records };
    },
    { entityName: "Events" },
    request,
  );
}
