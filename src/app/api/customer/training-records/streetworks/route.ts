import { withCustomerApi } from "@/lib/api/customerApi";
import { getCustomerStreetworksRecords } from "@/lib/services/customerTrainingRecordsService";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  return withCustomerApi(
    "GET /api/customer/training-records/streetworks",
    async (context) => {
      const records = await getCustomerStreetworksRecords(context.companyId);
      return { records };
    },
    { entityName: "Streetworks Training" },
    request,
  );
}
