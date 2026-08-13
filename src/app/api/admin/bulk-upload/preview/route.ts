import { withSharePointAdminApi, ValidationError } from "@/lib/api/adminApi";
import { logBulkUpload } from "@/lib/services/auditLogService";
import { previewBulkUpload } from "@/lib/services/bulkUpload/bulkUploadService";

export const dynamic = "force-dynamic";

export async function POST(request: Request) {
  return withSharePointAdminApi(
    "POST /api/admin/bulk-upload/preview",
    async (context, req) => {
      const contentType = req.headers.get("content-type") ?? "";
      if (!contentType.toLowerCase().includes("multipart/form-data")) {
        throw new ValidationError("Expected multipart form data.");
      }

      const form = await req.formData();
      const file = form.get("file");
      if (!(file instanceof File)) {
        throw new ValidationError("File is required.");
      }

      const importType = String(form.get("importType") ?? "").trim();
      const suppressRaw = form.get("suppressNotifications");
      const suppressNotifications =
        suppressRaw === null || suppressRaw === undefined
          ? true
          : !["false", "0", "no"].includes(String(suppressRaw).toLowerCase());
      const autoCreateRaw = form.get("autoCreateMissingCompanies");
      const autoCreateMissingCompanies =
        autoCreateRaw === null || autoCreateRaw === undefined
          ? true
          : !["false", "0", "no"].includes(
              String(autoCreateRaw).toLowerCase(),
            );
      const autoCreateDepartmentsRaw = form.get(
        "autoCreateMissingDepartments",
      );
      const autoCreateMissingDepartments =
        autoCreateDepartmentsRaw === null ||
        autoCreateDepartmentsRaw === undefined
          ? true
          : !["false", "0", "no"].includes(
              String(autoCreateDepartmentsRaw).toLowerCase(),
            );

      try {
        const result = await previewBulkUpload({
          importType,
          file,
          suppressNotifications,
          autoCreateMissingCompanies,
          autoCreateMissingDepartments,
        });

        await logBulkUpload({
          userEmail: context.loggedInEmail,
          phase: "preview",
          success: true,
          itemCount: result.summary.totalRows,
          request: req,
          metadata: {
            importType: result.importType,
            fileName: result.fileName,
            readyRows: result.summary.readyRows,
            warningRows: result.summary.warningRows,
            duplicateRows: result.summary.duplicateRows,
            errorRows: result.summary.errorRows,
            suppressNotifications: result.suppressNotifications,
            implemented: result.implemented,
          },
        });

        return result;
      } catch (error) {
        await logBulkUpload({
          userEmail: context.loggedInEmail,
          phase: "preview",
          success: false,
          request: req,
          errorMessage:
            error instanceof Error ? error.message : "Preview failed",
          metadata: {
            importType: importType || null,
            fileName: file.name || null,
          },
        });
        throw error;
      }
    },
    {
      errorMessage: "Failed to preview bulk upload",
      audit: false,
      entityName: "Bulk Upload",
    },
    request,
  );
}
