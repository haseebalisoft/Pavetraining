import { withCustomerApi } from "@/lib/api/customerApi";
import { getCustomerNporsRecords } from "@/lib/services/customerTrainingRecordsService";

export const dynamic = "force-dynamic";

/**
 * Customer NPORS training records.
 * Company is derived from Permissions List only — client companyId is rejected.
 */
export async function GET(request: Request) {
  return withCustomerApi(
    "GET /api/customer/training-records/npors",
    async (context) => {
      const records = await getCustomerNporsRecords(context.companyId, context);
      return { records };
    },
    { entityName: "NPORS Register" },
    request,
  );
}
