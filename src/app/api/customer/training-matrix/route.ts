import { withCustomerApi } from "@/lib/api/customerApi";
import { getCustomerMatrixRecords } from "@/lib/services/customerDashboardService";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  return withCustomerApi(
    "GET /api/customer/training-matrix",
    async (customer) => {
      const records = await getCustomerMatrixRecords(customer.companyName);
      return { companyName: customer.companyName, records };
    },
    { entityName: "Training Matrix", audit: true },
    request,
  );
}
