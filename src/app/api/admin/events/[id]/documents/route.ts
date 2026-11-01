import { withAdminApi, ValidationError } from "@/lib/api/adminApi";
import { writeAuditEvent } from "@/lib/services/auditLogService";
import { MAX_DOCUMENT_UPLOAD_BYTES } from "@/lib/services/customerDocumentUploadService";
import {
  listEventDocuments,
  uploadEventDocument,
} from "@/lib/services/eventDocumentService";

export const dynamic = "force-dynamic";

/** Documents attached to an event (visible and hidden — admin view). */
export async function GET(
  request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withAdminApi(
    "GET /api/admin/events/[id]/documents",
    async () => ({ documents: await listEventDocuments(id) }),
    { errorMessage: "Failed to load event documents" },
    request,
  );
}

/**
 * Attach a file to an event. Stored with the existing document storage
 * (Customer Documents → {company}/Company Documents/Events/{id}); opened by
 * customers through the existing, permission-checked document routes.
 * Form fields: file, customerVisible (default true).
 */
export async function POST(
  request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withAdminApi(
    "POST /api/admin/events/[id]/documents",
    async (adminContext, req) => {
      const contentType = req.headers.get("content-type") ?? "";
      if (!contentType.toLowerCase().includes("multipart/form-data")) {
        throw new ValidationError("Expected multipart form data.");
      }
      const form = await req.formData();
      const file = form.get("file");
      if (!(file instanceof File)) {
        throw new ValidationError("Choose a file to attach.");
      }
      if (file.size > MAX_DOCUMENT_UPLOAD_BYTES) {
        throw new ValidationError(
          `File exceeds the ${Math.floor(MAX_DOCUMENT_UPLOAD_BYTES / (1024 * 1024))} MB limit.`,
        );
      }
      const customerVisible = !["false", "0", "no"].includes(
        String(form.get("customerVisible") ?? "true").toLowerCase(),
      );
      const document = await uploadEventDocument({
        eventId: id,
        fileName: file.name,
        contentType: file.type || null,
        bytes: new Uint8Array(await file.arrayBuffer()),
        customerVisible,
        actorEmail: adminContext.loggedInEmail,
      });
      await writeAuditEvent({
        userEmail: adminContext.loggedInEmail,
        roleType: "Admin",
        action: "DOCUMENT_UPLOAD",
        entityType: "Event Documents",
        entityId: document.id,
        entityName: document.name,
        success: true,
        request: req,
        metadata: { eventId: id, customerVisible },
      });
      return { document };
    },
    { errorMessage: "Failed to attach the document" },
    request,
  );
}
