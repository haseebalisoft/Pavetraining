import { readFile } from "fs/promises";
import path from "path";

import { NextResponse } from "next/server";

import { withAdminApi, ValidationError } from "@/lib/api/adminApi";
import { listBulkUploadTemplates } from "@/lib/services/bulkUpload/bulkUploadService";
import {
  buildTemplateCsv,
  getBulkImportTemplate,
  getClientExcelTemplatePath,
} from "@/lib/services/bulkUpload/templates";
import type { BulkImportType } from "@/types/bulkUpload";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  const url = new URL(request.url);
  const download = url.searchParams.get("download")?.trim() ?? "";
  const format = (url.searchParams.get("format")?.trim() || "auto").toLowerCase();

  // Binary Excel templates for workforce / matrix — return file directly.
  if (download && format !== "csv") {
    const excelRelative = getClientExcelTemplatePath(download as BulkImportType);
    if (excelRelative) {
      return withAdminApi(
        "GET /api/admin/bulk-upload/templates?download=xlsx",
        async () => {
          const filePath = path.join(process.cwd(), "public", excelRelative);
          const bytes = await readFile(filePath);
          const template = getBulkImportTemplate(download as BulkImportType);
          return {
            __rawFile: true as const,
            fileName: template?.fileName ?? path.basename(excelRelative),
            contentType:
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            base64: bytes.toString("base64"),
          };
        },
        {
          errorMessage: "Failed to download template",
          audit: false,
        },
        request,
      ).then(async (response) => {
        // withAdminApi wraps JSON — unwrap raw file payload for binary download.
        try {
          const data = (await response.clone().json()) as {
            fileName?: string;
            contentType?: string;
            base64?: string;
            __rawFile?: boolean;
            error?: string;
          };
          if (data.__rawFile && data.base64 && data.fileName) {
            const buffer = Buffer.from(data.base64, "base64");
            return new NextResponse(buffer, {
              status: 200,
              headers: {
                "Content-Type":
                  data.contentType ||
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "Content-Disposition": `attachment; filename="${data.fileName}"`,
                "Cache-Control": "no-store",
              },
            });
          }
        } catch {
          // fall through to original JSON response
        }
        return response;
      });
    }
  }

  return withAdminApi(
    "GET /api/admin/bulk-upload/templates",
    async (_context, req) => {
      const reqUrl = new URL(req.url);
      const downloadType = reqUrl.searchParams.get("download")?.trim() ?? "";

      if (downloadType) {
        const template = getBulkImportTemplate(downloadType as BulkImportType);
        if (!template) {
          throw new ValidationError("Unknown template type.");
        }
        return {
          fileName: template.fileName.replace(/\.xlsx$/i, ".csv"),
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
