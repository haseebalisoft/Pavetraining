import { withAdminApi, ValidationError } from "@/lib/api/adminApi";
import {
  buildTemplateCsv,
  listBulkUploadTemplates,
} from "@/lib/services/bulkUpload/bulkUploadService";
import { getBulkImportTemplate } from "@/lib/services/bulkUpload/templates";
import type { BulkImportType } from "@/types/bulkUpload";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  return withAdminApi(
    "GET /api/admin/bulk-upload/templates",
    async (_context, req) => {
      const url = new URL(req.url);
      const download = url.searchParams.get("download")?.trim() ?? "";

      if (download) {
        const template = getBulkImportTemplate(download as BulkImportType);
        if (!template) {
          throw new ValidationError("Unknown template type.");
        }
        return {
          fileName: template.fileName,
          contentType: "text/csv; charset=utf-8",
          csv: buildTemplateCsv(template.importType),
        };
      }

      return {
        templates: listBulkUploadTemplates(),
      };
    },
    {
      errorMessage: "Failed to load bulk upload templates",
      audit: false,
    },
    request,
  );
}
