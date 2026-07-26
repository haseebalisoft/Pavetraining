import { withAdminApi, ValidationError } from "@/lib/api/adminApi";
import {
  MAX_DOCUMENT_UPLOAD_BYTES,
  uploadCustomerDocument,
} from "@/lib/services/customerDocumentUploadService";

export const dynamic = "force-dynamic";

export async function POST(request: Request) {
  return withAdminApi(
    "POST /api/admin/documents/upload",
    async (_ctx, req) => {
      const contentType = req.headers.get("content-type") ?? "";
      if (!contentType.toLowerCase().includes("multipart/form-data")) {
        throw new ValidationError("Expected multipart form data.");
      }

      const form = await req.formData();
      const file = form.get("file");
      if (!(file instanceof File)) {
        throw new ValidationError("File is required.");
      }
      if (file.size > MAX_DOCUMENT_UPLOAD_BYTES) {
        throw new ValidationError(
          `File exceeds the ${Math.floor(MAX_DOCUMENT_UPLOAD_BYTES / (1024 * 1024))} MB limit.`,
        );
      }

      const companyId = String(form.get("companyId") ?? "").trim();
      const candidateRaw = form.get("candidateId");
      const candidateId =
        candidateRaw === null || candidateRaw === undefined
          ? null
          : String(candidateRaw).trim() || null;
      const documentType = String(form.get("documentType") ?? "").trim();
      const customerVisibleRaw = String(
        form.get("customerVisible") ?? "true",
      ).toLowerCase();
      const customerVisible = !["false", "0", "no"].includes(customerVisibleRaw);

      const bytes = new Uint8Array(await file.arrayBuffer());
      const result = await uploadCustomerDocument({
        companyId,
        candidateId,
        documentType,
        customerVisible,
        fileName: file.name,
        contentType: file.type || null,
        bytes,
      });

      return {
        record: result.record,
        folderPath: result.folderPath,
        destinationFolder: result.destinationFolder,
      };
    },
    {
      errorMessage: "Failed to upload document",
      audit: true,
      entityName: "Customer Documents",
    },
    request,
  );
}
