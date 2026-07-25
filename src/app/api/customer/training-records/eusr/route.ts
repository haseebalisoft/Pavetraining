import { withCustomerApi } from "@/lib/api/customerApi";
import { getCustomerEusrRecords } from "@/lib/services/customerTrainingRecordsService";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  return withCustomerApi(
    "GET /api/customer/training-records/eusr",
    async (context) => {
      const records = await getCustomerEusrRecords(context.companyId);
      return { records };
    },
    { entityName: "EUSR Register" },
    request,
  );
}
