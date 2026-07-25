import { withCustomerApi } from "@/lib/api/customerApi";
import { getWorkforceByCompanyName } from "@/lib/services/workforceService";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  return withCustomerApi(
    "GET /api/customer/candidates",
    async (customer) => {
      const candidates = await getWorkforceByCompanyName(customer.companyName);
      return {
        companyName: customer.companyName,
        candidates: candidates.map((row) => ({
          id: row.id,
          candidateName: row.candidateName,
          companyName: row.companyName,
          workforceNumber: row.workforceNumber,
          department: row.department,
          status: row.status,
          trainingManager: row.trainingManager,
        })),
      };
    },
    { entityName: "Workforce List", audit: true },
    request,
  );
}
