import { withCustomerApi } from "@/lib/api/customerApi";

export const dynamic = "force-dynamic";

/**
 * Customer context from Permissions List.
 * Any companyId supplied by the client is rejected.
 */
export async function GET(request: Request) {
  return withCustomerApi(
    "GET /api/customer/context",
    async (context) => context,
    { entityName: "Customer Context" },
    request,
  );
}
