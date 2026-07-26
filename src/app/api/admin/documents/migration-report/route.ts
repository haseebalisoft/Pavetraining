import { withAdminApi } from "@/lib/api/adminApi";
import { buildCustomerDocumentsMigrationReport } from "@/lib/services/customerDocumentsMigrationReportService";

export const dynamic = "force-dynamic";

/**
 * Read-only report of documents whose library path does not match the
 * expected company/candidate/document-type folder structure.
 * Does not move files.
 */
export async function GET(request: Request) {
  return withAdminApi(
    "GET /api/admin/documents/migration-report",
    async () => ({
      report: await buildCustomerDocumentsMigrationReport(),
    }),
    { errorMessage: "Failed to build migration report" },
    request,
  );
}
