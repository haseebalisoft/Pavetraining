import { withCustomerApi } from "@/lib/api/customerApi";
import { getCustomerInHouseRecords } from "@/lib/services/customerTrainingRecordsService";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  return withCustomerApi(
    "GET /api/customer/training-records/in-house",
    async (context) => {
      const records = await getCustomerInHouseRecords(context.companyId);
      return { records };
    },
    { entityName: "In-House Certificates" },
    request,
  );
}
