import { withCustomerApi } from "@/lib/api/customerApi";
import { getCustomerDocumentRecords } from "@/lib/services/customerPortalService";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  return withCustomerApi(
    "GET /api/customer/documents",
    async (context) => {
      const records = await getCustomerDocumentRecords(
        context.companyId,
        context.canDownload,
      );
      return {
        canDownload: context.canDownload,
        records,
      };
    },
    { entityName: "Customer Documents" },
    request,
  );
}
