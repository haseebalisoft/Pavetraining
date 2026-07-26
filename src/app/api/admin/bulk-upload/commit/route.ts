import { withAdminApi, ValidationError } from "@/lib/api/adminApi";
import { logBulkUpload } from "@/lib/services/auditLogService";
import { commitBulkUpload } from "@/lib/services/bulkUpload/bulkUploadService";
import type { BulkCommitRowInput } from "@/types/bulkUpload";

export const dynamic = "force-dynamic";

export async function POST(request: Request) {
  return withAdminApi(
    "POST /api/admin/bulk-upload/commit",
    async (context, req) => {
      let body: Record<string, unknown>;
      try {
        body = (await req.json()) as Record<string, unknown>;
      } catch {
        throw new ValidationError("Expected JSON body.");
      }

      const importType = String(body.importType ?? "").trim();
      const fileName =
        body.fileName === null || body.fileName === undefined
          ? null
          : String(body.fileName);
      const duplicateMode =
        body.duplicateMode === null || body.duplicateMode === undefined
          ? "skip"
          : String(body.duplicateMode);
      const suppressNotifications =
        body.suppressNotifications === undefined
          ? true
          : Boolean(body.suppressNotifications);

      if (!Array.isArray(body.rows)) {
        throw new ValidationError("rows must be an array.");
      }

      const rows: BulkCommitRowInput[] = body.rows.map((raw, index) => {
        if (!raw || typeof raw !== "object") {
          throw new ValidationError(`Invalid row at index ${index}.`);
        }
        const row = raw as Record<string, unknown>;
        const rowNumber = Number(row.rowNumber);
        if (!Number.isFinite(rowNumber) || rowNumber < 1) {
          throw new ValidationError(`Invalid rowNumber at index ${index}.`);
        }
        const fieldsRaw = row.fields;
        if (!fieldsRaw || typeof fieldsRaw !== "object") {
          throw new ValidationError(`Missing fields at row ${rowNumber}.`);
        }
        const fields: Record<string, string | null> = {};
        for (const [key, value] of Object.entries(
          fieldsRaw as Record<string, unknown>,
        )) {
          if (value === null || value === undefined || value === "") {
            fields[key] = null;
          } else {
            fields[key] = String(value);
          }
        }
        return { rowNumber, fields };
      });

      try {
        const result = await commitBulkUpload({
          importType,
          fileName,
          duplicateMode,
          suppressNotifications,
          rows,
        });

        await logBulkUpload({
          userEmail: context.loggedInEmail,
          phase: "commit",
          success: true,
          itemCount: result.summary.importedRows,
          request: req,
          metadata: {
            importType: result.importType,
            fileName: result.fileName,
            duplicateMode: result.duplicateMode,
            suppressNotifications: result.suppressNotifications,
            totalRows: result.summary.totalRows,
            importedRows: result.summary.importedRows,
            skippedRows: result.summary.skippedRows,
            duplicateRows: result.summary.duplicateRows,
            errorRows: result.summary.errorRows,
            warningRows: result.summary.warningRows,
          },
        });

        return result;
      } catch (error) {
        await logBulkUpload({
          userEmail: context.loggedInEmail,
          phase: "commit",
          success: false,
          request: req,
          errorMessage:
            error instanceof Error ? error.message : "Commit failed",
          metadata: {
            importType: importType || null,
            fileName,
            duplicateMode,
            rowCount: rows.length,
          },
        });
        throw error;
      }
    },
    {
      errorMessage: "Failed to commit bulk upload",
      audit: false,
      entityName: "Bulk Upload",
    },
    request,
  );
}
