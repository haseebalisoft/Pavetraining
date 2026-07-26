import { withCustomerApi } from "@/lib/api/customerApi";
import { getAllowedWorkforceForCustomer } from "@/lib/services/customerAccessService";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  return withCustomerApi(
    "GET /api/customer/candidates",
    async (customer) => {
      const candidates = await getAllowedWorkforceForCustomer(customer);
      return {
        companyName: customer.companyName,
        customerRole: customer.customerRole,
        roleLabel: customer.roleLabel,
        accessScope: customer.normalizedAccessScope,
        candidates: candidates.map((row) => ({
          id: row.id,
          candidateName: row.candidateName,
          companyName: row.companyName,
          workforceNumber: row.workforceNumber,
          department: row.department,
          status: row.status,
          trainingManager: row.trainingManager,
          supervisor: row.supervisor,
        })),
      };
    },
    { entityName: "Workforce List", audit: true },
    request,
  );
}
