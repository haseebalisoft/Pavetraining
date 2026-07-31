import { withAdminApi, ValidationError } from "@/lib/api/adminApi";
import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import {
  MAX_IMAGE_BYTES,
  uploadAndSetListImage,
} from "@/lib/services/listThumbnailService";
import { getListItemByKey } from "@/lib/services/sharePointListService";
import { parseThumbnailField } from "@/lib/services/listThumbnailService";

export const dynamic = "force-dynamic";

export async function POST(
  request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withAdminApi(
    "POST /api/admin/workforce/[id]/photo",
    async (_ctx, req) => {
      const contentType = req.headers.get("content-type") ?? "";
      if (!contentType.toLowerCase().includes("multipart/form-data")) {
        throw new ValidationError("Expected multipart form data.");
      }
      const form = await req.formData();
      const file = form.get("file");
      if (!(file instanceof File)) {
        throw new ValidationError("Image file is required.");
      }
      if (file.size > MAX_IMAGE_BYTES) {
        throw new ValidationError("Image must be 10 MB or smaller.");
      }
      const bytes = new Uint8Array(await file.arrayBuffer());
      const fields = getSharePointFields("workforce");
      await uploadAndSetListImage({
        listKey: "workforce",
        itemId: id,
        fieldInternalName: fields.photo,
        fileName: file.name,
        bytes,
        contentType: file.type || null,
      });
      const item = await getListItemByKey("workforce", id);
      if (!item) throw new ValidationError("Candidate not found after upload.");
      const hasPhoto = Boolean(parseThumbnailField(item.fields[fields.photo]));
      return {
        id,
        photoUrl: hasPhoto ? `/api/media/workforce/${id}/photo` : null,
      };
    },
    { errorMessage: "Failed to upload candidate photo" },
    request,
  );
}
