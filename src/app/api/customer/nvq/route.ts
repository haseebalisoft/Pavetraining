import { withCustomerApi } from "@/lib/api/customerApi";
import { getCustomerNvqRecords } from "@/lib/services/customerPortalService";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  return withCustomerApi(
    "GET /api/customer/nvq",
    async (context) => {
      const records = await getCustomerNvqRecords(context.companyId);
      return { records };
    },
    { entityName: "NVQ Register" },
    request,
  );
}
