import { withCustomerApi } from "@/lib/api/customerApi";
import { getCustomerEventRecords } from "@/lib/services/customerPortalService";
import type { CustomerEventRecord } from "@/types/models";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  return withCustomerApi(
    "GET /api/customer/events",
    async (context) => {
      const records = await getCustomerEventRecords(context.companyId);
      // Security boundary: explicit public allowlist; internalNotes cannot reach customers.
      const publicRecords: CustomerEventRecord[] = records.map((record) => ({
        id: record.id,
        title: record.title,
        eventDate: record.eventDate,
        endDate: record.endDate,
        trainingAddress: record.trainingAddress,
        location: record.location,
        description: record.description,
        company: record.company,
      }));
      return { records: publicRecords };
    },
    { entityName: "Events" },
    request,
  );
}
