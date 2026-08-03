import { withCustomerApi } from "@/lib/api/customerApi";
import { submitCustomerSupportRequest } from "@/lib/services/customerSupportService";

export const dynamic = "force-dynamic";

export async function POST(request: Request) {
  return withCustomerApi(
    "POST /api/customer/support",
    async (customer, req) => {
      const body = (await (req ?? request).json()) as Record<string, unknown>;
      await submitCustomerSupportRequest(
        {
          name: String(body.name ?? ""),
          email: String(body.email ?? ""),
          phone: body.phone == null ? null : String(body.phone),
          companyName:
            body.companyName == null ? null : String(body.companyName),
          subject: String(body.subject ?? ""),
          message: String(body.message ?? ""),
        },
        customer,
      );
      return { ok: true };
    },
    { entityName: "Customer Support", audit: true },
    request,
  );
}
